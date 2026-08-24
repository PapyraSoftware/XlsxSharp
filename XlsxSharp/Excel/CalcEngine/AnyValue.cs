#nullable disable

using System.Globalization;
using XlsxSharp.Extensions;
using CollectionValue = XlsxSharp.Excel.CalcEngine.OneOf<
    XlsxSharp.Excel.CalcEngine.Array,
    XlsxSharp.Excel.CalcEngine.Reference
>;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// A discriminated union representing any value that can be passed around in the formula evaluation.
/// </summary>
internal readonly struct AnyValue
{
    private const int BlankValue = 0;
    private const int LogicalValue = 1;
    private const int NumberValue = 2;
    private const int TextValue = 3;
    private const int ErrorValue = 4;
    private const int ArrayValue = 5;
    private const int ReferenceValue = 6;

    private readonly byte _index;
    private readonly bool _logical;
    private readonly double _number;
    private readonly string _text;
    private readonly XLError _error;
    private readonly Array _array;
    private readonly Reference _reference;

    private AnyValue(
        byte index,
        bool logical,
        double number,
        string text,
        XLError error,
        Array array,
        Reference reference
    )
    {
        this._index = index;
        this._logical = logical;
        this._number = number;
        this._text = text;
        this._error = error;
        this._array = array;
        this._reference = reference;
    }

    /// <summary>
    /// A value of a blank cell or missing argument. Conversion methods mostly treat blank like 0 or an empty string.
    /// </summary>
    public static readonly AnyValue Blank = new(
        BlankValue,
        default,
        default,
        default,
        default,
        default,
        default
    );

    public static AnyValue From(bool logical) =>
        new(LogicalValue, logical, default, default, default, default, default);

    public static AnyValue From(double number) =>
        new(NumberValue, default, number, default, default, default, default);

    public static AnyValue From(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        return new AnyValue(TextValue, default, default, text, default, default, default);
    }

    public static AnyValue From(XLError error) =>
        new(ErrorValue, default, default, default, error, default, default);

    public static AnyValue From(Array array)
    {
        ArgumentNullException.ThrowIfNull(array);

        return new AnyValue(ArrayValue, default, default, default, default, array, default);
    }

    public static AnyValue From(Reference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);

        return new AnyValue(ReferenceValue, default, default, default, default, default, reference);
    }

    public static implicit operator AnyValue(bool logical) => From(logical);

    public static implicit operator AnyValue(double number) => From(number);

    public static implicit operator AnyValue(string text) => From(text);

    public static implicit operator AnyValue(XLError error) => From(error);

    public static implicit operator AnyValue(Array array) => From(array);

    public static implicit operator AnyValue(Reference reference) => From(reference);

    public bool IsBlank => this._index == BlankValue;

    public bool IsLogical => this._index == LogicalValue;

    public bool IsNumber => this._index == NumberValue;

    public bool IsText => this._index == TextValue;

    public bool IsError => this._index == ErrorValue;

    public bool IsArray => this._index == ArrayValue;

    public bool IsReference => this._index == ReferenceValue;

    /// <summary>
    /// Is the value a scalar (blank, logical, number, text or error).
    /// </summary>
    public bool IsScalarType =>
        this.IsBlank || this.IsLogical || this.IsNumber || this.IsText || this.IsError;

    public bool TryPickScalar(out ScalarValue scalar, out CollectionValue collection)
    {
        scalar = this._index switch
        {
            BlankValue => ScalarValue.Blank,
            LogicalValue => this._logical,
            NumberValue => this._number,
            TextValue => this._text,
            ErrorValue => this._error,
            _ => default,
        };
        collection = this._index switch
        {
            ArrayValue => this._array,
            ReferenceValue => this._reference,
            _ => default,
        };
        return this._index <= ErrorValue;
    }

    public bool TryPickError(out XLError error)
    {
        if (this._index == ErrorValue)
        {
            error = this._error;
            return true;
        }

        error = default;
        return false;
    }

    public bool TryPickArray(out Array array)
    {
        if (this._index == ArrayValue)
        {
            array = this._array;
            return true;
        }

        array = default;
        return false;
    }

    public bool TryPickReference(out Reference reference, out XLError error)
    {
        if (this._index == ReferenceValue)
        {
            reference = this._reference;
            error = default;
            return true;
        }

        reference = default;
        error = this._index == ErrorValue ? this._error : XLError.IncompatibleValue;
        return false;
    }

    /// <summary>
    /// Try to get a reference that is a one area from the value.
    /// </summary>
    /// <param name="area">The found area.</param>
    /// <param name="error">Original error, if the value is error, <c>#VALUE!</c> if type is not a reference or #REF! if more than one area in the reference.</param>
    /// <returns>True if area could be determined, false otherwise.</returns>
    public bool TryPickArea(out XLRangeAddress area, out XLError error)
    {
        if (this._index != ReferenceValue)
        {
            area = default;
            error = this._index == ErrorValue ? this._error : XLError.IncompatibleValue;
            return false;
        }

        if (this._reference.Areas.Count > 1)
        {
            area = default;
            error = XLError.CellReference;
            return false;
        }

        area = this._reference.Areas[0];
        error = default;
        return true;
    }

    /// <summary>
    /// Return array from a single area reference or array. If value is scalar, return false.
    /// </summary>
    public bool TryPickCollectionArray(out Array array, CalcContext ctx)
    {
        if (this.TryPickArea(out XLRangeAddress areaAddress, out _))
        {
            array = new ReferenceArray(areaAddress, ctx);
            return true;
        }

        if (this.TryPickArray(out array))
        {
            return true;
        }

        array = null;
        return false;
    }

    /// <summary>
    /// <para>
    /// Try to get a value more in line with an array formula semantic. The output is always
    /// either single value or an array.
    /// </para>
    /// <para>
    /// Single cell references are turned into a scalar, multi-area references are turned
    /// into <see cref="XLError.IncompatibleValue"/> and single-area references are turned
    /// into arrays.
    /// </para>
    /// </summary>
    /// <remarks>
    /// Note the difference in nomenclature: <em>single/multi value</em> vs <em>scalar/collection type</em>.
    /// </remarks>
    internal bool TryPickSingleOrMultiValue(
        out ScalarValue scalar,
        out Array array,
        CalcContext ctx
    )
    {
        if (this.TryPickScalar(out scalar, out CollectionValue collection))
        {
            array = default;
            return true;
        }

        // For some weird reason, 1x1 array doesn't count as a scalar, unlike single cell reference
        // proof {=TYPE(A1+1)} is 1 (scalar), but {=TYPE({1}+1)} is 64 (array).
        if (collection.TryPickT0(out array, out Reference reference))
        {
            scalar = default;
            return false;
        }

        if (reference.TryGetSingleCellValue(out scalar, ctx))
        {
            return true;
        }

        if (reference.Areas.Count > 1)
        {
            scalar = XLError.IncompatibleValue;
            return true;
        }

        array = new ReferenceArray(reference.Areas[0], ctx);
        return false;
    }

    public TResult Match<TResult>(
        Func<TResult> transformBlank,
        Func<bool, TResult> transformLogical,
        Func<double, TResult> transformNumber,
        Func<string, TResult> transformText,
        Func<XLError, TResult> transformError,
        Func<Array, TResult> transformArray,
        Func<Reference, TResult> transformReference
    ) =>
        this._index switch
        {
            BlankValue => transformBlank(),
            LogicalValue => transformLogical(this._logical),
            NumberValue => transformNumber(this._number),
            TextValue => transformText(this._text),
            ErrorValue => transformError(this._error),
            ArrayValue => transformArray(this._array),
            ReferenceValue => transformReference(this._reference),
            _ => throw new InvalidOperationException(),
        };

    #region Reference operators

    /// <summary>
    /// Implicit intersection for arguments of functions that don't accept range as a parameter (Excel 2016).
    /// </summary>
    /// <returns>Unchanged value for anything other than reference. Reference is changed into a single cell/#VALUE!</returns>
    public AnyValue ImplicitIntersection(CalcContext context) =>
        this.Match(
            () => Blank,
            logical => logical,
            number => number,
            text => text,
            logical => logical,
            array => array, // Array is unaffected by implicit intersection for operands - e.g. MMULT(COS({0,0});COS({0;0})) = 2
            reference =>
            {
                if (reference.IsSingleCell())
                {
                    return reference;
                }

                return reference
                    .ImplicitIntersection(context.FormulaAddress)
                    .Match<AnyValue>(singleCellReference => singleCellReference, error => error);
            }
        );

    /// <summary>
    /// Create a new reference that has one area that contains both operands or #VALUE! if not possible.
    /// </summary>
    public static AnyValue ReferenceRange(in AnyValue left, in AnyValue right, CalcContext ctx)
    {
        OneOf<Reference, XLError> leftConversionResult = ConvertToReference(left);
        if (!leftConversionResult.TryPickT0(out Reference leftReference, out XLError leftError))
        {
            return leftError;
        }

        OneOf<Reference, XLError> rightConversionResult = ConvertToReference(right);
        if (!rightConversionResult.TryPickT0(out Reference rightReference, out XLError rightError))
        {
            return rightError;
        }

        return Reference
            .RangeOp(leftReference, rightReference, ctx.Worksheet)
            .Match<AnyValue>(reference => reference, error => error);
    }

    /// <summary>
    /// Create a new reference by combining areas of both arguments. Areas of the new reference can overlap = some overlapping
    /// cells might be counted multiple times (<c>SUM((A1;A1)) = 2</c> if <c>A1</c> is <c>1</c>).
    /// </summary>
    public static AnyValue ReferenceUnion(in AnyValue left, in AnyValue right)
    {
        OneOf<Reference, XLError> leftConversionResult = ConvertToReference(left);
        if (!leftConversionResult.TryPickT0(out Reference leftReference, out XLError leftError))
        {
            return leftError;
        }

        OneOf<Reference, XLError> rightConversionResult = ConvertToReference(right);
        if (!rightConversionResult.TryPickT0(out Reference rightReference, out XLError rightError))
        {
            return rightError;
        }

        return Reference.UnionOp(leftReference, rightReference);
    }

    private static OneOf<Reference, XLError> ConvertToReference(in AnyValue value) =>
        value._index switch
        {
            ReferenceValue => value._reference,
            ErrorValue => value._error,
            _ => XLError.IncompatibleValue,
        };

    #endregion

    #region Arithmetic unary operations

    public AnyValue UnaryPlus() =>
        // Unary plus doesn't even convert to number. Type and value is retained.
        this;

    public AnyValue UnaryMinus(CalcContext context) => UnaryOperation(this, x => -x, context);

    public AnyValue UnaryPercent(CalcContext context) =>
        UnaryOperation(this, x => x / 100.0, context);

    private static AnyValue UnaryOperation(
        in AnyValue value,
        Func<double, double> operatorFn,
        CalcContext context
    )
    {
        bool isSingle = value.TryPickSingleOrMultiValue(
            out ScalarValue single,
            out Array array,
            context
        );
        if (isSingle)
        {
            return UnaryArithmeticOp(single, operatorFn, context).ToAnyValue();
        }

        return array.Apply(arrayConst => UnaryArithmeticOp(arrayConst, operatorFn, context));
    }

    private static ScalarValue UnaryArithmeticOp(
        ScalarValue value,
        Func<double, double> op,
        CalcContext ctx
    )
    {
        OneOf<double, XLError> conversionResult = value.ToNumber(ctx.Culture);
        if (!conversionResult.TryPickT0(out double number, out XLError error))
        {
            return error;
        }

        return op(number);
    }

    #endregion

    #region Arithmetic binary operators

    public static AnyValue BinaryPlus(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(
                    in leftItem,
                    in rightItem,
                    static (lhs, rhs) => lhs + rhs,
                    ctx
                );
            },
            context
        );

    public static AnyValue BinaryMinus(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(
                    in leftItem,
                    in rightItem,
                    static (lhs, rhs) => lhs - rhs,
                    ctx
                );
            },
            context
        );

    public static AnyValue BinaryMult(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(
                    in leftItem,
                    in rightItem,
                    static (lhs, rhs) => lhs * rhs,
                    ctx
                );
            },
            context
        );

    public static AnyValue BinaryDiv(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(
                    in leftItem,
                    in rightItem,
                    static (lhs, rhs) => rhs == 0.0 ? XLError.DivisionByZero : lhs / rhs,
                    ctx
                );
            },
            context
        );

    public static AnyValue BinaryExp(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return BinaryArithmeticOp(
                    in leftItem,
                    in rightItem,
                    static (lhs, rhs) =>
                        lhs == 0 && rhs == 0 ? XLError.NumberInvalid : Math.Pow(lhs, rhs),
                    ctx
                );
            },
            context
        );

    #endregion

    #region Comparison operators

    public static AnyValue IsEqual(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp == 0, static error => error);
            },
            context
        );

    public static AnyValue IsNotEqual(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp != 0, static error => error);
            },
            context
        );

    public static AnyValue IsGreaterThan(
        in AnyValue left,
        in AnyValue right,
        CalcContext context
    ) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp > 0, static error => error);
            },
            context
        );

    public static AnyValue IsGreaterThanOrEqual(
        in AnyValue left,
        in AnyValue right,
        CalcContext context
    ) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp >= 0, static error => error);
            },
            context
        );

    public static AnyValue IsLessThan(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp < 0, static error => error);
            },
            context
        );

    public static AnyValue IsLessThanOrEqual(
        in AnyValue left,
        in AnyValue right,
        CalcContext context
    ) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                return CompareValues(leftItem, rightItem, ctx.Culture)
                    .Match<ScalarValue>(static cmp => cmp <= 0, static error => error);
            },
            context
        );

    #endregion

    public static AnyValue Concat(in AnyValue left, in AnyValue right, CalcContext context) =>
        BinaryOperation(
            in left,
            in right,
            static (in ScalarValue leftItem, in ScalarValue rightItem, CalcContext ctx) =>
            {
                OneOf<string, XLError> leftTextResult = leftItem.ToText(ctx.Culture);
                if (!leftTextResult.TryPickT0(out string leftText, out XLError leftError))
                {
                    return leftError;
                }

                OneOf<string, XLError> rightTextResult = rightItem.ToText(ctx.Culture);
                if (!rightTextResult.TryPickT0(out string rightText, out XLError rightError))
                {
                    return rightError;
                }

                return leftText + rightText;
            },
            context
        );

    private static AnyValue BinaryOperation(
        in AnyValue left,
        in AnyValue right,
        BinaryFunc func,
        CalcContext context
    )
    {
        bool isLeftSingle = left.TryPickSingleOrMultiValue(
            out ScalarValue leftSingle,
            out Array leftArray,
            context
        );
        bool isRightSingle = right.TryPickSingleOrMultiValue(
            out ScalarValue rightSingle,
            out Array rightArray,
            context
        );

        if (isLeftSingle && isRightSingle)
        {
            return func(in leftSingle, in rightSingle, context).ToAnyValue();
        }

        if (isLeftSingle)
        {
            ScalarArray broadcastedLeftArray = new(leftSingle, rightArray.Width, rightArray.Height);
            return broadcastedLeftArray.Apply(rightArray, func, context);
        }

        if (isRightSingle)
        {
            ScalarArray broadcastedRightArray = new(rightSingle, leftArray.Width, leftArray.Height);
            return leftArray.Apply(broadcastedRightArray, func, context);
        }

        int unifiedRows = Math.Max(leftArray.Height, rightArray.Height);
        int unifiedColumns = Math.Max(leftArray.Width, rightArray.Width);

        Array leftBroadcastedArray = leftArray.Broadcast(unifiedRows, unifiedColumns);
        Array rightBroadcastedArray = rightArray.Broadcast(unifiedRows, unifiedColumns);

        return leftBroadcastedArray.Apply(rightBroadcastedArray, func, context);
    }

    private static ScalarValue BinaryArithmeticOp(
        in ScalarValue left,
        in ScalarValue right,
        BinaryNumberFunc func,
        CalcContext ctx
    )
    {
        OneOf<double, XLError> leftConversionResult = left.ToNumber(ctx.Culture);
        if (!leftConversionResult.TryPickT0(out double leftNumber, out XLError leftError))
        {
            return leftError;
        }

        OneOf<double, XLError> rightConversionResult = right.ToNumber(ctx.Culture);
        if (!rightConversionResult.TryPickT0(out double rightNumber, out XLError rightError))
        {
            return rightError;
        }

        return func(leftNumber, rightNumber).Match<ScalarValue>(number => number, error => error);
    }

    /// <summary>
    /// Compare two scalar values using Excel semantic. Rules for comparison are following:
    /// <list type="bullet">
    ///     <item>Logical is always greater than any text (thus transitively greater than any number)</item>
    ///     <item>Text is always greater than any number, even if empty string</item>
    ///     <item>Logical are compared by value</item>
    ///     <item>Numbers are compared by value</item>
    ///     <item>Text is compared by through case insensitive comparison for workbook culture.</item>
    ///     <item>
    ///         If any argument is error, return error (general rule for all operators).
    ///         If all args are errors, pick leftmost error (technically it is left to
    ///         implementation, but excel sems to be using left one)
    ///     </item>
    /// </list>
    /// </summary>
    /// <param name="left">Left hand operand of the comparison.</param>
    /// <param name="right">Right hand operand of the comparison.</param>
    /// <param name="culture">Culture to use for comparison.</param>
    /// <returns>
    ///     Return -1 (negative)  if left less than right
    ///     Return 1 (positive) if left greater than left
    ///     Return 0 if both operands are considered equal.
    /// </returns>
    private static OneOf<int, XLError> CompareValues(
        ScalarValue left,
        ScalarValue right,
        CultureInfo culture
    ) =>
        left.Match(
            culture,
            _ =>
                right.Match<OneOf<int, XLError>, CultureInfo>(
                    culture,
                    _ => 0,
                    (rightLogical, _) => false.CompareTo(rightLogical),
                    (rightNumber, _) => 0.0.CompareTo(rightNumber),
                    (rightText, culture) =>
                        string.Compare(string.Empty, rightText, culture, CompareOptions.IgnoreCase),
                    (rightError, _) => rightError
                ),
            (leftLogical, _) =>
                right.Match<OneOf<int, XLError>, bool>(
                    leftLogical,
                    leftLogical => leftLogical.CompareTo(false),
                    (rightLogical, leftLogical) => leftLogical.CompareTo(rightLogical),
                    (rightNumber, _) => 1,
                    (rightText, _) => 1,
                    (rightError, _) => rightError
                ),
            (leftNumber, _) =>
                right.Match<OneOf<int, XLError>, double>(
                    leftNumber,
                    leftNumber => leftNumber.CompareTo(0.0),
                    (rightLogical, _) => -1,
                    (rightNumber, leftNumber) => leftNumber.CompareTo(rightNumber),
                    (rightText, _) => -1,
                    (rightError, _) => rightError
                ),
            (leftText, culture) =>
                right.Match<OneOf<int, XLError>, string, CultureInfo>(
                    leftText,
                    culture,
                    (leftText, culture) =>
                        string.Compare(leftText, string.Empty, culture, CompareOptions.IgnoreCase),
                    (rightLogical, _, _) => -1,
                    (rightNumber, _, _) => 1,
                    (rightText, leftText, culture) =>
                        string.Compare(leftText, rightText, culture, CompareOptions.IgnoreCase),
                    (rightError, _, _) => rightError
                ),
            (leftError, _) => leftError
        );

    public override string ToString() =>
        this._index switch
        {
            BlankValue => "Blank",
            LogicalValue => $"Logical: {this._logical.ToString().ToUpper()}",
            NumberValue => $"Number: {this._number}",
            TextValue => $"Text: {this._text}",
            ErrorValue => $"Error: {this._error.ToDisplayString()}",
            ArrayValue => $"Array{this._array.Height}x{this._array.Width}",
            ReferenceValue =>
                $"Reference: {string.Join(",", this._reference.Areas.Select(a => $"{a.FirstAddress}:{a.LastAddress}"))}",
            _ => throw new InvalidOperationException(),
        };

    /// <summary>
    /// Get 2d size of the value. For scalars, it's 1x1, for multi-area references,
    /// it's also 1x1,because it is converted to <c>#VALUE!</c> error.
    /// </summary>
    public (int Rows, int Columns) GetArraySize()
    {
        if (this.IsScalarType)
        {
            return (1, 1);
        }

        if (this.TryPickArray(out Array array))
        {
            return (array.Height, array.Width);
        }

        if (this.TryPickArea(out XLRangeAddress area, out _))
        {
            return (area.RowSpan, area.ColumnSpan);
        }

        // Multi area is just error = scalar
        return (1, 1);
    }

    /// <summary>
    /// Return the array value.
    /// </summary>
    /// <exception cref="InvalidCastException" />
    public Array GetArray() =>
        this._index == ArrayValue ? this._array : throw new InvalidCastException();

    private delegate OneOf<double, XLError> BinaryNumberFunc(double lhs, double rhs);
}

internal delegate ScalarValue BinaryFunc(in ScalarValue lhs, in ScalarValue rhs, CalcContext ctx);

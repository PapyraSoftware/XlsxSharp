using System.Globalization;
using XlsxSharp.Excel.CalcEngine.Exceptions;
using XlsxSharp.Excel.CalcEngine.Functions;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Parser;
using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Excel.CalcEngine;

internal sealed class CalcContext
{
    private readonly XLCalcEngine _calcEngine;
    private readonly XLWorkbook? _workbook;
    private readonly XLWorksheet? _worksheet;
    private readonly IXLAddress? _formulaAddress;
    private readonly bool _recursive;

    public CalcContext(XLCalcEngine calcEngine, CultureInfo culture, XLCell cell)
        : this(calcEngine, culture, cell.Worksheet.Workbook, cell.Worksheet, cell.Address) { }

    public CalcContext(
        XLCalcEngine calcEngine,
        CultureInfo culture,
        XLWorkbook? workbook,
        XLWorksheet? worksheet,
        IXLAddress? formulaAddress,
        bool recursive = false
    )
    {
        this._calcEngine = calcEngine;
        this._workbook = workbook;
        this._worksheet = worksheet;
        this._formulaAddress = formulaAddress;
        this._recursive = recursive;
        this.Culture = culture;
        this.CancellationToken = workbook?.CancellationToken ?? CancellationToken.None;
    }

    // LEGACY: Remove once legacy functions are migrated
    internal XLCalcEngine CalcEngine => this._calcEngine ?? throw new MissingContextException();

    /// <summary>
    /// Worksheet of the cell the formula is calculating.
    /// </summary>
    public XLWorkbook Workbook => this._workbook ?? throw new MissingContextException();

    /// <summary>
    /// Worksheet of the cell the formula is calculating.
    /// </summary>
    public XLWorksheet Worksheet => this._worksheet ?? throw new MissingContextException();

    /// <summary>
    /// Address of the calculated formula.
    /// </summary>
    public IXLAddress FormulaAddress => this._formulaAddress ?? throw new MissingContextException();

    /// <summary>
    /// A culture used for comparisons and conversions (e.g. text to number).
    /// </summary>
    public CultureInfo Culture { get; }

    /// <summary>
    /// Excel 2016 and earlier doesn't support dynamic array formulas (it used an array formulas instead). As a consequence,
    /// all arguments for scalar functions where passed through implicit intersection before calling the function.
    /// </summary>
    public static bool UseImplicitIntersection => true;

    /// <summary>
    /// Should functions be calculated per item of multi-values argument in the scalar parameters.
    /// </summary>
    public bool IsArrayCalculation { get; set; }

    /// <summary>
    /// Sheet that is being recalculated. If set, formula can read dirty
    /// values from other sheets, but not from this sheet.
    /// </summary>
    public string? RecalculateSheetName { get; set; }

    internal Point FormulaPoint =>
        new(this.FormulaAddress.RowNumber, this.FormulaAddress.ColumnNumber);

    /// <summary>
    /// What date system should be used in calculation. Either 1900 or 1904.
    /// </summary>
    internal bool Use1904DateSystem { get; init; } = false;

    /// <summary>
    /// An upper limit (exclusive) of used calendar system.
    /// </summary>
    internal double DateSystemUpperLimit =>
        this.Use1904DateSystem
            ? XlsxSharp.XLHelper.Calendar1904UpperLimit
            : XlsxSharp.XLHelper.Calendar1900UpperLimit;

    internal CancellationToken CancellationToken { get; }

    /// <summary>
    /// A helper method to check is user cancelled the calculation in function loops.
    /// </summary>
    internal void ThrowIfCancelled() => this.CancellationToken.ThrowIfCancellationRequested();

    internal ScalarValue GetCellValue(XLWorksheet? sheet, int rowNumber, int columnNumber)
    {
        sheet ??= this.Worksheet;
        ValueSlice valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        Point point = new(rowNumber, columnNumber);
        XLCellFormula? formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);

        if (formula is null)
        {
            return valueSlice.GetCellValue(point);
        }

        if (!formula.IsDirty)
        {
            return valueSlice.GetCellValue(point);
        }

        // Used when only one sheet should be recalculated, leaving other sheets with their data.
        if (
            this.RecalculateSheetName is not null
            && !XlsxSharp.XLHelper.SheetComparer.Equals(sheet.Name, this.RecalculateSheetName)
        )
        {
            return valueSlice.GetCellValue(point);
        }

        // A special branch for functions out of cells (e.g. worksheet.Evaluate("A1+2")).
        // These functions are not a part of calculation chain and thus reordering a chain
        // for them doesn't make sense.
        if (this._recursive)
        {
            XLCell? cell = sheet.GetCell(point);
            return cell?.Value ?? Blank.Value;
        }

        throw new GettingDataException(
            new SheetPoint(sheet.Name, new Point(rowNumber, columnNumber))
        );
    }

    /// <summary>
    /// This method goes over slices and returns a value for each non-blank cell. Because it is using
    /// slice iterators, it scales with number of cells, not a size of area in reference (i.e. it works
    /// fine even if reference is <c>A1:XFD1048576</c>). It also works for 3D references.
    /// </summary>
    internal IEnumerable<ScalarValue> GetNonBlankValues(Reference reference)
    {
        foreach (XLRangeAddress area in reference.Areas)
        {
            XLWorksheet sheet = area.Worksheet ?? this.Worksheet;
            Area range = Area.FromRangeAddress(area);

            // A value can be either in a non-empty value slice or a empty cell with a formula.
            XLCellsCollection.SlicesEnumerator enumerator =
                sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
            while (enumerator.MoveNext())
            {
                Point point = enumerator.Current;
                ScalarValue scalarValue = this.GetCellValue(sheet, point.Row, point.Column);
                if (!scalarValue.IsBlank)
                {
                    yield return scalarValue;
                }
            }
        }
    }

    /// <summary>
    /// Return all points in the <paramref name="areaReference" /> that satisfy the <paramref name="criteria" />.
    /// </summary>
    internal IEnumerable<Point> GetCriteriaPoints(XLRangeAddress areaReference, Criteria criteria)
    {
        XLWorksheet sheet = areaReference.Worksheet ?? this.Worksheet;
        Area area = Area.FromRangeAddress(areaReference);

        // This is a performance optimization when user specifies a whole column
        // in the tally function (e.g. SUMIF(A:B, "5", C:D)).
        if (criteria.CanBlankValueMatch)
        {
            // Criteria can match blank cells, thus it's not possible to use optimized
            // used enumerators and we have to check value of each cell.
            foreach (Point point in area)
            {
                ScalarValue scalarValue = this.GetCellValue(sheet, point.Row, point.Column);
                if (criteria.Match(scalarValue))
                {
                    yield return point;
                }
            }
        }
        else
        {
            // The criteria can never match blank cells. That means we can skip all blank
            // cells entirely and use optimized used enumerators.
            XLCellsCollection.SlicesEnumerator enumerator =
                sheet.Internals.CellsCollection.ForValuesAndFormulas(area);
            while (enumerator.MoveNext())
            {
                Point point = enumerator.Current;
                ScalarValue scalarValue = this.GetCellValue(sheet, point.Row, point.Column);
                if (criteria.Match(scalarValue))
                {
                    yield return point;
                }
            }
        }
    }

    internal IEnumerable<ScalarValue> GetFilteredNonBlankValues(
        Reference reference,
        string function,
        bool skipHiddenRows = false
    )
    {
        // Allocate one per call, because visitor holds info whether function was found in a formula.
        FunctionVisitor visitor = new(function);
        foreach (XLRangeAddress area in reference.Areas)
        {
            XLWorksheet sheet = area.Worksheet ?? this.Worksheet;
            Area range = Area.FromRangeAddress(area);
            int currentRow = 0;
            bool rowIsHidden = true;

            // A value can be either in a non-empty value slice or a empty cell with a formula.
            XLCellsCollection.SlicesEnumerator enumerator =
                sheet.Internals.CellsCollection.ForValuesAndFormulas(range);
            while (enumerator.MoveNext())
            {
                Point point = enumerator.Current;

                if (skipHiddenRows)
                {
                    // If row changed, update hidden info about current row
                    if (currentRow != point.Row)
                    {
                        currentRow = point.Row;
                        rowIsHidden =
                            sheet.Internals.RowsCollection.TryGetValue(currentRow, out XLRow? row)
                            && row.IsHidden;
                    }

                    if (rowIsHidden)
                    {
                        continue;
                    }
                }

                XLCellFormula? formula = sheet.Internals.CellsCollection.FormulaSlice.Get(point);
                if (CallsFunction(formula, visitor))
                {
                    continue;
                }

                ScalarValue scalarValue = this.GetCellValue(sheet, point.Row, point.Column);
                if (!scalarValue.IsBlank)
                {
                    yield return scalarValue;
                }
            }
        }

        yield break;

        static bool CallsFunction(XLCellFormula? formula, FunctionVisitor visitor)
        {
            if (formula is null)
            {
                return false;
            }

            if (!formula.A1.Contains(visitor.FunctionName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            ParserFactory.Create(visitor).ParseFormula(formula.A1, visitor);
            if (!visitor.Found)
            {
                return false;
            }

            // In order to reuse same visitor without allocation, clear the found flag.
            visitor.Clear();
            return true;
        }
    }

    /// <summary>
    /// This method should be used mostly for range arguments. If a value is scalar,
    /// return a single value enumerable.
    /// </summary>
    internal IEnumerable<ScalarValue> GetNonBlankValues(AnyValue value)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
        {
            if (scalar.IsBlank)
            {
                return System.Array.Empty<ScalarValue>();
            }

            return new ScalarArray(scalar, 1, 1);
        }

        if (collection.TryPickT0(out Array? array, out Reference? reference))
        {
            return array.Where(x => !x.IsBlank);
        }

        return this.GetNonBlankValues(reference);
    }

    internal IEnumerable<ScalarValue> GetAllValues(AnyValue value)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
        {
            return new ScalarArray(scalar, 1, 1);
        }

        if (collection.TryPickT0(out Array? array, out Reference? reference))
        {
            return array;
        }

        return this.GetAllCellValues(reference);
    }

    internal IEnumerable<ScalarValue> GetAllCellValues(Reference reference)
    {
        foreach (XLRangeAddress area in reference.Areas)
        {
            XLWorksheet? sheet = area.Worksheet;
            foreach (Point point in Area.FromRangeAddress(area))
            {
                yield return this.GetCellValue(sheet, point.Row, point.Column);
            }
        }
    }

    private class FunctionVisitor : CollectVisitor<FunctionVisitor>
    {
        public FunctionVisitor(string function) => this.FunctionName = function;

        internal string FunctionName { get; }

        public bool Found { get; private set; }

        public void Clear() => this.Found = false;

        public override object? Function(
            FunctionVisitor context,
            SymbolRange range,
            ReadOnlySpan<char> functionName,
            IReadOnlyList<object?> arguments
        )
        {
            this.Found =
                this.Found
                || functionName.Equals(
                    this.FunctionName.AsSpan(),
                    StringComparison.OrdinalIgnoreCase
                );
            return default;
        }
    }
}

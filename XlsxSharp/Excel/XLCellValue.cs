#nullable disable

using System;
using System.Diagnostics;
using System.Globalization;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

/// <summary>
/// A value of a single cell. It contains a value a specific <see cref="Type"/>.
/// Structure provides following group of methods:
/// <list type="bullet">
///   <item><c>Is*</c> properties to check type (<see cref="IsNumber"/>, <see cref="IsBlank"/>...)</item>
///   <item><c>Get*</c> methods that return the stored value or throw <see cref="InvalidCastException"/> for incorrect type.</item>
///   <item>Explicit operators to convert <c>XLCellValue</c> to a concrete type. It is an equivalent of <c>Get*</c> methods.</item>
///   <item><c>TryConvert</c> methods to try to get value of a specific type, even if the value is of a different type.</item>
/// </list>
/// </summary>
[DebuggerDisplay("{Type} {_text != null ? (object)_text : (object)_value}")]
public readonly struct XLCellValue
    : IEquatable<XLCellValue>,
        IEquatable<Blank>,
        IEquatable<Boolean>,
        IEquatable<Double>,
        IEquatable<String>,
        IEquatable<XLError>,
        IEquatable<DateTime>,
        IEquatable<TimeSpan>,
        IEquatable<int>
{
    private readonly double _value;
    private readonly string _text;

    private XLCellValue(Blank _)
        : this()
    {
        this.Type = XLDataType.Blank;
    }

    private XLCellValue(bool logical)
        : this()
    {
        this.Type = XLDataType.Boolean;
        this._value = logical ? 1d : 0d;
    }

    private XLCellValue(double number)
        : this()
    {
        if (Double.IsNaN(number) || Double.IsInfinity(number))
        {
            throw new ArgumentException("Value can't be NaN or infinity.", nameof(number));
        }

        this.Type = XLDataType.Number;
        this._value = number;
    }

    private XLCellValue(string text)
        : this()
    {
        ArgumentNullException.ThrowIfNull(text);

        if (text.Length > 32767)
        {
            throw new ArgumentOutOfRangeException(
                nameof(text),
                "Cells can hold a maximum of 32,767 characters."
            );
        }

        this.Type = XLDataType.Text;
        this._text = text;
    }

    private XLCellValue(XLError error)
        : this()
    {
        if (error < XLError.NullValue || error > XLError.NoValueAvailable)
        {
            throw new ArgumentOutOfRangeException(nameof(error));
        }

        this.Type = XLDataType.Error;
        this._value = (double)error;
    }

    private XLCellValue(DateTime dateTime)
        : this()
    {
        this.Type = XLDataType.DateTime;
        this._value = dateTime.ToSerialDateTime();
    }

    private XLCellValue(TimeSpan timeSpan)
        : this()
    {
        this.Type = XLDataType.TimeSpan;
        this._value = timeSpan.ToSerialDateTime();
    }

    private XLCellValue(XLDataType type, double value)
        : this()
    {
        if (Double.IsNaN(value) || Double.IsInfinity(value))
        {
            throw new ArgumentException("Value can't be NaN or infinity.", nameof(value));
        }

        this.Type = type;
        this._value = value;
    }

    /// <summary>
    /// Type of the value.
    /// </summary>
    public XLDataType Type { get; }

    /// <summary>
    /// Is the type of value <c>Blank</c>?
    /// </summary>
    public bool IsBlank => this.Type == XLDataType.Blank;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.Boolean"/>?
    /// </summary>
    public bool IsBoolean => this.Type == XLDataType.Boolean;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.Number"/>?
    /// </summary>
    public bool IsNumber => this.Type == XLDataType.Number;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.Text"/>?
    /// </summary>
    public bool IsText => this.Type == XLDataType.Text;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.Error"/>?
    /// </summary>
    public bool IsError => this.Type == XLDataType.Error;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.DateTime"/>?
    /// </summary>
    public bool IsDateTime => this.Type == XLDataType.DateTime;

    /// <summary>
    /// Is the type of value <see cref="XLDataType.TimeSpan"/>?
    /// </summary>
    public bool IsTimeSpan => this.Type == XLDataType.TimeSpan;

    /// <summary>
    /// Is the value <see cref="XLDataType.Number"/> or <see cref="XLDataType.DateTime"/> or <see cref="XLDataType.TimeSpan"/>.
    /// </summary>
    public bool IsUnifiedNumber => this.IsNumber || this.IsDateTime || this.IsTimeSpan;

    public static implicit operator XLCellValue(Blank blank) => new(blank);

    public static implicit operator XLCellValue(bool logical) => new(logical);

    public static implicit operator XLCellValue(string text) =>
        text is not null ? new XLCellValue(text) : new XLCellValue(Blank.Value);

    public static implicit operator XLCellValue(XLError error) => new(error);

    public static implicit operator XLCellValue(DateTime dateTime) => new(dateTime);

    public static implicit operator XLCellValue(TimeSpan timeSpan) => new(timeSpan);

    public static implicit operator XLCellValue(sbyte number) => new(number);

    public static implicit operator XLCellValue(byte number) => new(number);

    public static implicit operator XLCellValue(short number) => new(number);

    public static implicit operator XLCellValue(ushort number) => new(number);

    public static implicit operator XLCellValue(int number) => new(number);

    public static implicit operator XLCellValue(uint number) => new(number);

    public static implicit operator XLCellValue(long number) => new(number);

    public static implicit operator XLCellValue(ulong number) => new(number);

    public static implicit operator XLCellValue(float number) => new(number);

    public static implicit operator XLCellValue(double number) => new(number);

    public static implicit operator XLCellValue(decimal number) => new(decimal.ToDouble(number));

    public static implicit operator XLCellValue(sbyte? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(byte? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(short? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(ushort? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(int? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(uint? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(long? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(ulong? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(float? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(double? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(decimal? numberOrBlank) =>
        numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(DateTime? dateTimeOrBlank) =>
        dateTimeOrBlank.HasValue ? dateTimeOrBlank.Value : Blank.Value;

    public static implicit operator XLCellValue(TimeSpan? timeSpanOrBlank) =>
        timeSpanOrBlank.HasValue ? timeSpanOrBlank.Value : Blank.Value;

    /// <summary>
    /// Creates an <see cref="XLCellValue"/> from an <see cref="object"/>. If the type of the object has an implicit conversion operator then it is used.
    /// Otherwise, the <see cref="Convert.ToString(object, IFormatProvider)"/> method is used to convert the provided object to a string.
    /// <para/>
    /// The following types and their nullable counterparts are supported without requiring to be converted to a string:
    /// <list type="bullet">
    ///   <item><see cref="Blank"/></item>
    ///   <item><see cref="bool"/></item>
    ///   <item><see cref="string"/></item>
    ///   <item><see cref="XLError"/></item>
    ///   <item><see cref="DateTime"/></item>
    ///   <item><see cref="TimeSpan"/></item>
    ///   <item><see cref="sbyte"/></item>
    ///   <item><see cref="byte"/></item>
    ///   <item><see cref="short"/></item>
    ///   <item><see cref="ushort"/></item>
    ///   <item><see cref="int"/></item>
    ///   <item><see cref="uint"/></item>
    ///   <item><see cref="long"/></item>
    ///   <item><see cref="ulong"/></item>
    ///   <item><see cref="float"/></item>
    ///   <item><see cref="double"/></item>
    ///   <item><see cref="decimal"/></item>
    /// </list>
    /// </summary>
    /// <param name="obj">The object to convert.</param>
    /// <param name="provider">An object that supplies culture-specific formatting information.</param>
    /// <returns>An <see cref="XLCellValue"/> representation of the object.</returns>
    public static XLCellValue FromObject(object obj, IFormatProvider provider = null)
    {
        return obj switch
        {
            null => Blank.Value,
            Blank blank => blank,
            bool logical => logical,
            string text => text,
            XLError error => error,
            DateTime dateTime => dateTime,
            TimeSpan timeSpan => timeSpan,
            sbyte number => number,
            byte number => number,
            short number => number,
            ushort number => number,
            int number => number,
            uint number => number,
            long number => number,
            ulong number => number,
            float number => number,
            double number => number,
            decimal number => number,
            _ => Convert.ToString(obj, provider),
        };
    }

    /// <summary>
    /// A function used during data insertion to convert an <c>object</c> to <c>XLCellValue</c>.
    /// </summary>
    internal static XLCellValue FromInsertedObject(object value)
    {
        XLCellValue convertedValue = value switch
        {
            null => Blank.Value,
            Blank blankValue => blankValue,
            Boolean logical => logical,
            SByte number => number,
            Byte number => number,
            Int16 number => number,
            UInt16 number => number,
            Int32 number => number,
            UInt32 number => number,
            Int64 number => number,
            UInt64 number => number,
            Single number => number,
            Double number => number,
            Decimal number => number,
            String text => text,
            XLError error => error,
            DateTime date => date,
            DateTimeOffset dateOfs => dateOfs.DateTime,
            TimeSpan timeSpan => timeSpan,
            _ => value.ToString(), // Other things, like chars ect are just turned to string
        };
        return convertedValue;
    }

    /// <summary>
    /// Try to convert a string into a string value, doing your best. If no other type can be
    /// extracted, consider it a text.
    /// </summary>
    /// <param name="text">Text to parse into a value.</param>
    /// <param name="culture">Culture used to parse numbers.</param>
    /// <returns>Parsed value.</returns>
    internal static XLCellValue FromText(string text, CultureInfo culture)
    {
        // AutoFilter custom filter operand can be stored as `1 1/2` and Excel correctly
        // interprets it as a `1.5`. Same for 2015-01-01, therefore use `TextToNumber` that
        // should deal with any weird formats.
        if (text is null)
        {
            return Blank.Value;
        }

        if (text == String.Empty)
        {
            return Blank.Value;
        }

        if (StringComparer.OrdinalIgnoreCase.Equals("TRUE", text))
        {
            return true;
        }

        if (StringComparer.OrdinalIgnoreCase.Equals("FALSE", text))
        {
            return false;
        }

        if (ScalarValue.TextToNumber(text, culture).TryPickT0(out double number, out _))
        {
            return number;
        }

        if (XLErrorParser.TryParseError(text, out XLError error))
        {
            return error;
        }

        return text;
    }

    /// <inheritdoc cref="GetBlank"/>
    public static explicit operator Blank(XLCellValue value) => value.GetBlank();

    /// <inheritdoc cref="GetBoolean"/>
    public static explicit operator Boolean(XLCellValue value) => value.GetBoolean();

    /// <inheritdoc cref="GetNumber"/>
    public static explicit operator Double(XLCellValue value) => value.GetNumber();

    /// <inheritdoc cref="GetText"/>
    public static explicit operator String(XLCellValue value) => value.GetText();

    /// <inheritdoc cref="GetError"/>
    public static explicit operator XLError(XLCellValue value) => value.GetError();

    /// <inheritdoc cref="GetDateTime"/>
    public static implicit operator DateTime(XLCellValue value) => value.GetDateTime();

    /// <inheritdoc cref="GetTimeSpan"/>
    public static implicit operator TimeSpan(XLCellValue value) => value.GetTimeSpan();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.Blank"/>,
    /// return <see cref="Blank.Value"/>, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Blank GetBlank() => this.IsBlank ? Blank.Value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.Boolean"/>,
    /// return logical, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Boolean GetBoolean() =>
        this.IsBoolean ? this._value != 0d : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.Number"/>,
    /// return number, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Double GetNumber() => this.IsNumber ? this._value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.Text"/>,
    /// return text, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public String GetText() => this.IsText ? this._text : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.Error"/>,
    /// return error, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public XLError GetError() =>
        this.IsError ? (XLError)this._value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.DateTime"/>,
    /// return date time, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public DateTime GetDateTime() =>
        this.IsDateTime ? this._value.ToSerialDateTime() : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XLDataType.TimeSpan"/>,
    /// return time span, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public TimeSpan GetTimeSpan() =>
        this.IsTimeSpan ? this._value.ToSerialTimeSpan() : throw new InvalidCastException();

    internal static XLCellValue FromSerialDateTime(double serialDateTime) =>
        new(XLDataType.DateTime, serialDateTime);

    internal static XLCellValue FromSerialTimeSpan(double serialTimeSpan) =>
        new(XLDataType.TimeSpan, serialTimeSpan);

    /// <summary>
    /// Get a number, either directly from type number or from serialized time span
    /// or serialized date time.
    /// </summary>
    /// <exception cref="InvalidCastException">If type is not <see cref="XLDataType.Number"/> or
    /// <see cref="XLDataType.DateTime"/> or <see cref="XLDataType.TimeSpan"/>.</exception>
    public double GetUnifiedNumber()
    {
        if (this.IsUnifiedNumber)
        {
            return this._value;
        }

        throw new InvalidCastException("Value is not a number.");
    }

    internal Object ToObject()
    {
        return this.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Boolean => this.GetBoolean(),
            XLDataType.Number => this.GetNumber(),
            XLDataType.Text => this.GetText(),
            XLDataType.Error => this.GetError(),
            XLDataType.DateTime => this.GetDateTime(),
            XLDataType.TimeSpan => this.GetTimeSpan(),
            _ => throw new InvalidCastException(),
        };
    }

    /// <summary>
    /// Return text representation of a value in current culture.
    /// </summary>
    public override string ToString() => this.ToString(CultureInfo.CurrentCulture);

    /// <summary>
    /// Return text representation of a value in the specified culture.
    /// </summary>
    public string ToString(CultureInfo culture) =>
        this.Type switch
        {
            XLDataType.Blank => string.Empty,
            XLDataType.Boolean => this.GetBoolean() ? "TRUE" : "FALSE",
            XLDataType.Number => this._value.ToString(culture),
            XLDataType.Text => this._text,
            XLDataType.Error => this.GetError().ToDisplayString(),
            XLDataType.DateTime => this.GetDateTime().ToString(culture),
            XLDataType.TimeSpan => this.GetTimeSpan().ToExcelString(culture),
            _ => throw new InvalidOperationException(),
        };

    public bool Equals(XLCellValue other)
    {
        return this.Type == other.Type
            && this._value.Equals(other._value)
            && this._text == other._text;
    }

    public bool Equals(Blank other)
    {
        return this.IsBlank;
    }

    public bool Equals(bool other)
    {
        return this.IsBoolean && this.GetBoolean() == other;
    }

    public bool Equals(double other)
    {
        return this.IsNumber && this._value.Equals(other);
    }

    /// <summary>
    /// Is the cell value text and is equal to the <paramref name="other"/>?
    /// Text comparison is case sensitive.
    /// </summary>
    public bool Equals(string other)
    {
        return this.IsText && this._text == other;
    }

    public bool Equals(XLError other)
    {
        return this.IsError && this.GetError() == other;
    }

    public bool Equals(DateTime other)
    {
        return this.IsDateTime && this.GetDateTime() == other;
    }

    public bool Equals(TimeSpan other)
    {
        return this.IsTimeSpan && this.GetTimeSpan() == other;
    }

    public bool Equals(int other)
    {
        return this.Equals((double)other);
    }

    public override bool Equals(object obj)
    {
        return obj is XLCellValue other && this.Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            int hashCode = this._value.GetHashCode();
            hashCode = (hashCode * 397) ^ (this._text != null ? this._text.GetHashCode() : 0);
            hashCode = (hashCode * 397) ^ (int)this.Type;
            return hashCode;
        }
    }

    /// <summary>
    /// Get a value, if it is a <see cref="XLDataType.Text"/>.
    /// </summary>
    /// <returns>True if value was retrieved, false otherwise.</returns>
    public bool TryGetText(out string value)
    {
        if (this.IsText)
        {
            value = this._text;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XLDataType.Blank"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XLDataType.Blank"/>.</item>
    ///   <item>Type <see cref="XLDataType.Text"/> and the text is empty.</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out Blank value)
    {
        bool isBlankLike = this.IsBlank || (this.IsText && this.GetText().Length == 0);
        if (isBlankLike)
        {
            value = Blank.Value;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XLDataType.Boolean"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XLDataType.Boolean"/>.</item>
    ///   <item>Type <see cref="XLDataType.Number"/>, then the value of <c>0</c> means <c>false</c> and any other value is <c>true</c>.</item>
    ///   <item>Type <see cref="XLDataType.Text"/> and the value is <c>TRUE</c> or <c>FALSE</c> (case insensitive). Note that calc engine
    ///   generally doesn't coerce text to logical (e.g. arithmetic operations), though it happens in most function arguments (e.g.
    ///   <c>IF</c> or <c>AND</c>).</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out Boolean value)
    {
        switch (this.Type)
        {
            case XLDataType.Boolean:
                value = this.GetBoolean();
                return true;
            case XLDataType.Number:
                value = this.GetNumber() != 0;
                return true;
            case XLDataType.Text
                when String.Equals(this.GetText(), "TRUE", StringComparison.OrdinalIgnoreCase):
                value = true;
                return true;
            case XLDataType.Text
                when String.Equals(this.GetText(), "FALSE", StringComparison.OrdinalIgnoreCase):
                value = false;
                return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XLDataType.Number"/> and return it.
    /// <list type="bullet">
    ///   <item>Double value - return the value.</item>
    ///   <item>Boolean value - return the <c>0</c> for <c>TRUE</c> and <c>1</c> for <c>FALSE</c>.</item>
    ///   <item>Text value - use the <c>VALUE</c> semantic to convert a text to a number.</item>
    ///   <item>DateTime value - return the serial date time number.</item>
    ///   <item>TimeSpan value - return the serial time span value.</item>
    /// </list>
    /// </summary>
    /// <remarks>Note that the coercion is current culture specific (e.g. decimal separators can differ).</remarks>
    /// <param name="value">The converted value. Result is never <c>infinity</c> or <c>NaN</c>.</param>
    /// <param name="culture">The culture used to convert the value for texts.</param>
    public bool TryConvert(out Double value, CultureInfo culture)
    {
        switch (this.Type)
        {
            case XLDataType.Number:
            case XLDataType.DateTime:
            case XLDataType.TimeSpan:
                value = this.GetUnifiedNumber();
                return true;
            case XLDataType.Boolean:
                value = this.GetBoolean() ? 1 : 0;
                return true;
            case XLDataType.Text:
            {
                OneOf<double, XLError> coercionResult = ScalarValue.TextToNumber(
                    this.GetText(),
                    culture
                );
                if (coercionResult.TryPickT0(out double number, out _))
                {
                    value = number;
                    return true;
                }

                break;
            }
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="DateTime"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XLDataType.DateTime"/>.</item>
    ///   <item>Type <see cref="XLDataType.Number"/> and when the number is interpreted is a serial date time, it falls within the DateTime range.</item>
    ///   <item>Type <see cref="XLDataType.TimeSpan"/> and when the number is interpreted is a serial date time, it falls within the DateTime range.</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out DateTime value)
    {
        if (this.IsUnifiedNumber)
        {
            double number = this.GetUnifiedNumber();
            if (number.IsValidOADateNumber())
            {
                value = number.ToSerialDateTime();
                return true;
            }
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="TimeSpan"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XLDataType.TimeSpan"/>.</item>
    ///   <item>Type <see cref="XLDataType.Number"/>, the number is interpreted is a time span date time.</item>
    ///   <item>Type <see cref="XLDataType.Text"/> and it can be parsed as a time span (accepts even hours over 24 hours).</item>
    /// </list>
    /// </summary>
    /// <param name="value">The converted value.</param>
    /// <param name="culture">The culture used to get time and decimal separators.</param>
    public bool TryConvert(out TimeSpan value, CultureInfo culture)
    {
        if (this.IsTimeSpan)
        {
            value = this.GetTimeSpan();
            return true;
        }

        if (this.IsNumber)
        {
            value = this.GetUnifiedNumber().ToSerialTimeSpan();
            return true;
        }

        if (this.IsText && TimeSpanParser.TryParseTime(this.GetText(), culture, out TimeSpan ts))
        {
            value = ts;
            return true;
        }

        value = default;
        return false;
    }
}

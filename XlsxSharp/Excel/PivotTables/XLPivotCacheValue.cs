using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Excel;

/// <summary>
/// Represents a single value in a pivot cache record.
/// </summary>
internal readonly struct XLPivotCacheValue
{
    /// <summary>
    /// A memory used to hold value of a <see cref="Type"/>. Its
    /// interpretation depends on the type. It doesn't hold value
    /// for strings directly, because GC doesn't allow aliasing
    /// same 8 bytes for number or references. For strings, it contains
    /// an index to a string storage array that is stored separately.
    /// </summary>
    private readonly double _value;

    private XLPivotCacheValue(XLPivotCacheValueType type, double value)
    {
        this.Type = type;
        this._value = value;
    }

    internal XLPivotCacheValueType Type { get; }

    internal static XLPivotCacheValue ForMissing() => new(XLPivotCacheValueType.Missing, 0);

    internal static XLPivotCacheValue ForNumber(double number)
    {
        if (double.IsNaN(number) || double.IsInfinity(number))
        {
            throw new ArgumentOutOfRangeException();
        }

        return new XLPivotCacheValue(XLPivotCacheValueType.Number, number);
    }

    internal static XLPivotCacheValue ForBoolean(bool boolean) =>
        new(XLPivotCacheValueType.Boolean, boolean ? 1 : 0);

    internal static XLPivotCacheValue ForError(XLError error) =>
        new(XLPivotCacheValueType.Error, (int)error);

    internal static XLPivotCacheValue ForText(string text, List<string> storage)
    {
        int index = storage.Count;
        storage.Add(text);
        return new XLPivotCacheValue(
            XLPivotCacheValueType.String,
            BitConverter.Int64BitsToDouble(index)
        );
    }

    internal static XLPivotCacheValue ForText(
        string text,
        Dictionary<string, int> stringMap,
        List<string> storage
    )
    {
        if (!stringMap.TryGetValue(text, out int index))
        {
            index = storage.Count;
            storage.Add(text);
            stringMap.Add(text, index);
            return new XLPivotCacheValue(
                XLPivotCacheValueType.String,
                BitConverter.Int64BitsToDouble(index)
            );
        }

        return new XLPivotCacheValue(
            XLPivotCacheValueType.String,
            BitConverter.Int64BitsToDouble(index)
        );
    }

    internal static XLPivotCacheValue ForDateTime(DateTime dateTime) =>
        new(XLPivotCacheValueType.DateTime, BitConverter.Int64BitsToDouble(dateTime.Ticks));

    internal static XLPivotCacheValue ForIndex(uint index) =>
        new(XLPivotCacheValueType.Index, BitConverter.Int64BitsToDouble(index));

    internal XLCellValue GetCellValue(
        List<string> stringStorage,
        XLPivotCacheSharedItems sharedItems
    )
    {
        switch (this.Type)
        {
            case XLPivotCacheValueType.Missing:
                return Blank.Value;

            case XLPivotCacheValueType.Number:
                return this._value;

            case XLPivotCacheValueType.Boolean:
                return this._value != 0;

            case XLPivotCacheValueType.Error:
                return (XLError)this._value;

            case XLPivotCacheValueType.String:
                return this.GetText(stringStorage);

            case XLPivotCacheValueType.DateTime:
                return this.GetDateTime();

            case XLPivotCacheValueType.Index:
                uint intIndex = unchecked((uint)BitConverter.DoubleToInt64Bits(this._value));
                return sharedItems[intIndex];

            default:
                throw new NotSupportedException();
        }
    }

    internal double GetNumber() => this._value;

    internal bool GetBoolean() => this._value != 0;

    internal XLError GetError() => (XLError)this._value;

    internal string GetText(IReadOnlyList<string> stringStorage)
    {
        int stringIndex = unchecked((int)BitConverter.DoubleToInt64Bits(this._value));
        return stringStorage[stringIndex];
    }

    internal DateTime GetDateTime()
    {
        long ticks = BitConverter.DoubleToInt64Bits(this._value);
        return new DateTime(ticks);
    }

    internal uint GetIndex() => unchecked((uint)BitConverter.DoubleToInt64Bits(this._value));
}

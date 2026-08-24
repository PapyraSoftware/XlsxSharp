using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Excel;

/// <summary>
/// All values of a cache field for a pivot table.
/// </summary>
internal class XLPivotCacheValues
{
    private readonly XLPivotCacheSharedItems _sharedItems;

    private readonly List<XLPivotCacheValue> _values;

    private readonly List<string> _stringStorage;

    private bool _containsBlank;

    private bool _containsNumber;

    private double? _minValue;

    private double? _maxValue;

    /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsInteger"/>
    private bool _containsInteger;

    /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsString"/>
    private bool _containsString;

    /// <inheritdoc cref="XLPivotCacheValuesStats.LongText"/>
    private bool _longText;

    /// <inheritdoc cref="XLPivotCacheValuesStats.ContainsDate"/>
    private bool _containsDate;

    private long? _minDateTicks;

    private long? _maxDateTicks;

    internal XLPivotCacheValues(ValueSlice valueSlice, int column, Area area)
    {
        this._sharedItems = new XLPivotCacheSharedItems();
        this._values = [];
        this._stringStorage = [];

        this.Initialize(valueSlice, column, area);
    }

    internal XLPivotCacheValues(XLPivotCacheSharedItems sharedItems, XLPivotCacheValuesStats stats)
    {
        this._sharedItems = sharedItems;
        this._values = [];
        this._stringStorage = [];

        // Have a separate fields instead of one large struct. That way,
        // the flags are more easily set when record values are being added.
        this._containsBlank = stats.ContainsBlank;
        this._containsNumber = stats.ContainsNumber;
        this._containsInteger = stats.ContainsInteger;
        this._minValue = stats.MinValue;
        this._maxValue = stats.MaxValue;
        this._containsString = stats.ContainsString;
        this._longText = stats.LongText;
        this._containsDate = stats.ContainsDate;
        this._minDateTicks = stats.MinDate?.Ticks;
        this._maxDateTicks = stats.MaxDate?.Ticks;
    }

    internal XLPivotCacheValuesStats Stats
    {
        get
        {
            DateTime? minDate =
                this._containsDate && this._minDateTicks is not null
                    ? new DateTime(this._minDateTicks.Value)
                    : null;
            DateTime? maxDate =
                this._containsDate && this._maxDateTicks is not null
                    ? new DateTime(this._maxDateTicks.Value)
                    : null;

            return new XLPivotCacheValuesStats(
                this._containsBlank,
                this._containsNumber,
                this._containsInteger,
                this._minValue,
                this._maxValue,
                this._containsString,
                this._longText,
                this._containsDate,
                minDate,
                maxDate
            );
        }
    }

    internal int Count => this._values.Count;

    internal int SharedCount => this._sharedItems.Count;

    internal XLPivotCacheSharedItems SharedItems => this._sharedItems;

    internal void AddMissing()
    {
        this._values.Add(XLPivotCacheValue.ForMissing());
        this._containsBlank = true;
    }

    internal void AddNumber(double number)
    {
        this._values.Add(XLPivotCacheValue.ForNumber(number));
        this.AdjustStats(number);
    }

    internal void AddBoolean(bool boolean)
    {
        this._values.Add(XLPivotCacheValue.ForBoolean(boolean));

        // [MS-OI29500]: In Office, boolean and error are considered strings in the context of the containsString attribute.
        this._containsString = true;
    }

    internal void AddError(XLError error)
    {
        this._values.Add(XLPivotCacheValue.ForError(error));

        // [MS-OI29500]: In Office, boolean and error are considered strings in the context of the containsString attribute.
        this._containsString = true;
    }

    internal void AddString(string text)
    {
        this._values.Add(XLPivotCacheValue.ForText(text, this._stringStorage));
        this.AdjustStats(text);
    }

    internal void AddDateTime(DateTime dateTime)
    {
        this._values.Add(XLPivotCacheValue.ForDateTime(dateTime));
        this.AdjustStats(dateTime);
    }

    internal void AddIndex(uint index)
    {
        if (index >= this._sharedItems.Count)
        {
            throw new ArgumentException("Index is referencing non-existent shared item.");
        }

        this._values.Add(XLPivotCacheValue.ForIndex(index));

        // Get value referenced by added index value, so stats can be updated.
        XLPivotCacheValue cacheValue = this._sharedItems.GetValue(index);
        switch (cacheValue.Type)
        {
            case XLPivotCacheValueType.Missing:
                this._containsBlank = true;
                break;
            case XLPivotCacheValueType.Number:
                this.AdjustStats(cacheValue.GetNumber());
                break;
            case XLPivotCacheValueType.Boolean:
                this._containsString = true;
                break;
            case XLPivotCacheValueType.Error:
                this._containsString = true;
                break;
            case XLPivotCacheValueType.String:
                this.AdjustStats(this._sharedItems.GetStringValue(index));
                break;
            case XLPivotCacheValueType.DateTime:
                this.AdjustStats(cacheValue.GetDateTime());
                break;
            default:
                throw new NotSupportedException();
        }
    }

    internal XLPivotCacheValue GetValue(int recordIdx) => this._values[recordIdx];

    internal string GetText(XLPivotCacheValue value)
    {
        Debug.Assert(value.Type == XLPivotCacheValueType.String);
        return value.GetText(this._stringStorage);
    }

    internal void AllocateCapacity(int recordCount) => this._values.Capacity = recordCount;

    internal IEnumerable<XLCellValue> GetCellValues()
    {
        foreach (XLPivotCacheValue value in this._values)
        {
            yield return value.GetCellValue(this._stringStorage, this._sharedItems);
        }
    }

    /// <summary>
    /// Get or add a value to the shared items. Throw, if value is not in items.
    /// </summary>
    /// <returns>Index in shared items.</returns>
    internal int GetOrAddSharedItem(XLCellValue value)
    {
        int sharedItemsIndex = this._sharedItems.IndexOf(value);
        if (sharedItemsIndex >= 0)
        {
            return sharedItemsIndex;
        }

        // Not in shared items, make sure it actually is present.
        if (!this.ContainsRecord(value))
        {
            throw new ArgumentException($"Value '{value}' not among cache field values.");
        }

        this._sharedItems.Add(value);

        return this._sharedItems.Count - 1;
    }

    /// <summary>
    /// Is among the <c>value</c> among values of the record.
    /// </summary>
    private bool ContainsRecord(XLCellValue value)
    {
        for (int index = 0; index < this._values.Count; ++index)
        {
            XLPivotCacheValue recordValue = this.GetValue(index);
            XLCellValue cacheValue = recordValue.GetCellValue(
                this._stringStorage,
                this._sharedItems
            );
            if (XLCellValueComparer.OrdinalIgnoreCase.Equals(cacheValue, value))
            {
                return true;
            }
        }

        return false;
    }

    private void Initialize(ValueSlice valueSlice, int column, Area area)
    {
        HashSet<XLCellValue> uniqueItems = new(XLCellValueComparer.OrdinalIgnoreCase);
        for (int row = area.TopRow + 1; row <= area.BottomRow; ++row)
        {
            XLCellValue value = valueSlice.GetCellValue(new Point(row, column));

            // Add to shared items first, because value can be an index to shared items.
            if (uniqueItems.Add(value))
            {
                this._sharedItems.Add(value);
            }

            switch (value.Type)
            {
                case XLDataType.Blank:
                    this.AddMissing();
                    break;
                case XLDataType.Boolean:
                    this.AddBoolean(value.GetBoolean());
                    break;
                case XLDataType.Number:
                    this.AddNumber(value.GetNumber());
                    break;
                case XLDataType.Text:
                    this.AddString(value.GetText());
                    break;
                case XLDataType.Error:
                    this.AddError(value.GetError());
                    break;
                case XLDataType.DateTime:
                    this.AddDateTime(value.GetDateTime());
                    break;
                case XLDataType.TimeSpan:
                    // TimeSpan is represented as datetime in pivot cache, e.g. 14:30 into 1899-12-30T14:30:00
                    DateTime adjustedTimeSpan = DateTime.FromOADate(0).Add(value.GetTimeSpan());
                    this.AddDateTime(adjustedTimeSpan);
                    break;
                default:
                    throw new UnreachableException();
            }
        }
    }

    [SuppressMessage(
        "ReSharper",
        "CompareOfFloatsByEqualityOperator",
        Justification = "double.IsInteger() in NET7 uses same method."
    )]
    private void AdjustStats(double number)
    {
        // containsInt is true only if all numbers are integers.
        this._containsInteger =
            // First ever number is an integer.
            (!this._containsNumber && number == Math.Truncate(number))
            ||
            // Subsequent number is an integer.
            (this._containsInteger && number == Math.Truncate(number));
        this._containsNumber = true;
        this._minValue = this._minValue is null ? number : Math.Min(this._minValue.Value, number);
        this._maxValue = this._maxValue is null ? number : Math.Max(this._maxValue.Value, number);
    }

    private void AdjustStats(string text)
    {
        this._containsString = true;
        this._longText = this._longText || text.Length > 255;
    }

    private void AdjustStats(DateTime dateTime)
    {
        this._containsDate = true;
        long dateTicks = dateTime.Ticks;
        this._minDateTicks = this._minDateTicks is null
            ? dateTicks
            : Math.Min(this._minDateTicks.Value, dateTicks);
        this._maxDateTicks = this._maxDateTicks is null
            ? dateTicks
            : Math.Max(this._maxDateTicks.Value, dateTicks);
    }
}

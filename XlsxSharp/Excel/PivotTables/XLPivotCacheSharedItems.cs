using System;
using System.Collections.Generic;
using System.Diagnostics;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

/// <summary>
/// <para>
/// A list of <see cref="XLPivotCacheValue"/> in the pivot table cache
/// definition. Generally, it contains all strings of the field records
/// (record just indexes them through <see cref="XLPivotCacheValueType.Index"/>)
/// and also values used directly in pivot table (e.g. filter field reference
/// the table definition, not record).
/// </para>
/// <para>
/// Shared items can't contain <see cref="XLPivotCacheValueType.Index"/>.
/// </para>
/// </summary>
internal class XLPivotCacheSharedItems
{
    private readonly List<XLPivotCacheValue> _values = [];

    /// <summary>
    /// Storage of strings to save 8 bytes per <c>XLPivotCacheValue</c>
    /// (reference can't be aliased with a number).
    /// </summary>
    private readonly List<string> _stringStorage = [];

    /// <summary>
    /// Strings in a pivot table are case-insensitive.
    /// </summary>
    private readonly Dictionary<string, int> _stringMap = new(StringComparer.OrdinalIgnoreCase);

    internal XLCellValue this[uint index] =>
        this.GetValue(index).GetCellValue(this._stringStorage, this);

    internal int Count => this._values.Count;

    internal void Add(XLCellValue value)
    {
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
                DateTime timeSpan = value.GetTimeSpan().ToSerialDateTime().ToSerialDateTime();
                this.AddDateTime(timeSpan);
                break;
            default:
                throw new UnreachableException();
        }
    }

    internal void AddMissing()
    {
        this._values.Add(XLPivotCacheValue.ForMissing());
    }

    internal void AddNumber(double number)
    {
        this._values.Add(XLPivotCacheValue.ForNumber(number));
    }

    internal void AddBoolean(bool boolean)
    {
        this._values.Add(XLPivotCacheValue.ForBoolean(boolean));
    }

    internal void AddError(XLError error)
    {
        this._values.Add(XLPivotCacheValue.ForError(error));
    }

    internal void AddString(string text)
    {
        // Shared items doesn't distinguish between two texts that differ only in case.
        if (!this._stringMap.ContainsKey(text))
        {
            int index = this._stringStorage.Count;
            this._values.Add(XLPivotCacheValue.ForText(text, this._stringStorage));
            this._stringMap.Add(text, index);
        }
    }

    internal void AddDateTime(DateTime dateTime)
    {
        this._values.Add(XLPivotCacheValue.ForDateTime(dateTime));
    }

    internal IEnumerable<XLCellValue> GetCellValues()
    {
        foreach (XLPivotCacheValue value in this._values)
        {
            yield return value.GetCellValue(this._stringStorage, this);
        }
    }

    internal XLPivotCacheValue GetValue(uint index)
    {
        return this._values[checked((int)index)];
    }

    internal string GetStringValue(uint index)
    {
        XLPivotCacheValue value = this.GetValue(index);
        return value.GetText(this._stringStorage);
    }

    /// <summary>
    /// Get index of value or -1 if not among shared items.
    /// </summary>
    internal int IndexOf(XLCellValue value)
    {
        for (int index = 0; index < this._values.Count; ++index)
        {
            XLPivotCacheValue sharedValue = this._values[index];
            XLCellValue cacheValue = sharedValue.GetCellValue(this._stringStorage, this);
            if (XLCellValueComparer.OrdinalIgnoreCase.Equals(cacheValue, value))
            {
                return index;
            }
        }

        return -1;
    }
}

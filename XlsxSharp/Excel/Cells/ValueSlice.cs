using System;
using System.Collections.Generic;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Excel;

/// <summary>
/// A slice of a single worksheet for values of a cell.
/// </summary>
internal class ValueSlice : ISlice
{
    private readonly Slice<XLValueSliceContent> _values = new();
    private readonly SharedStringTable _sst;

    internal ValueSlice(SharedStringTable sst) => this._sst = sst;

    public bool IsEmpty => this._values.IsEmpty;

    public int MaxColumn => this._values.MaxColumn;

    public int MaxRow => this._values.MaxRow;

    public Dictionary<int, int>.KeyCollection UsedColumns => this._values.UsedColumns;

    public IEnumerable<int> UsedRows => this._values.UsedRows;

    public void Clear(Area area)
    {
        this.DereferenceTextInRange(area);
        this._values.Clear(area);
    }

    public void DeleteAreaAndShiftLeft(Area areaToDelete)
    {
        this.DereferenceTextInRange(areaToDelete);
        this._values.DeleteAreaAndShiftLeft(areaToDelete);
    }

    public void DeleteAreaAndShiftUp(Area areaToDelete)
    {
        this.DereferenceTextInRange(areaToDelete);
        this._values.DeleteAreaAndShiftUp(areaToDelete);
    }

    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false) =>
        this._values.GetEnumerator(area, reverse);

    public void InsertAreaAndShiftDown(Area areaToInsert)
    {
        // Only pushed out references have to be dereferenced, other text references just move.
        if (areaToInsert.BottomRow < XlsxSharp.XLHelper.MaxRowNumber)
        {
            Area belowRange = areaToInsert.BelowRange();
            int pushedOutRows = Math.Min(areaToInsert.Height, belowRange.Height);
            Area pushedOutRange = belowRange.SliceFromBottom(pushedOutRows);
            this.DereferenceTextInRange(pushedOutRange);
        }

        this._values.InsertAreaAndShiftDown(areaToInsert);
    }

    public void InsertAreaAndShiftRight(Area areaToInsert)
    {
        // Only pushed out references have to be dereferenced, other text references just move.
        if (areaToInsert.RightColumn < XlsxSharp.XLHelper.MaxColumnNumber)
        {
            Area rightRange = areaToInsert.RightRange();
            int pushedOutColumns = Math.Min(areaToInsert.Width, rightRange.Width);
            Area pushedOutRange = rightRange.SliceFromRight(pushedOutColumns);
            this.DereferenceTextInRange(pushedOutRange);
        }

        this._values.InsertAreaAndShiftRight(areaToInsert);
    }

    public bool IsUsed(Point address) => this._values.IsUsed(address);

    public void Swap(Point sp1, Point sp2) => this._values.Swap(sp1, sp2);

    internal XLCellValue GetCellValue(Point point)
    {
        ref readonly XLValueSliceContent cellValue = ref this._values[point];
        XLDataType type = cellValue.Type;
        double value = cellValue.Value;
        return type switch
        {
            XLDataType.Blank => Blank.Value,
            XLDataType.Boolean => value != 0,
            XLDataType.Number => value,
            XLDataType.Text => this._sst[(int)value],
            XLDataType.Error => (XLError)value,
            XLDataType.DateTime => XLCellValue.FromSerialDateTime(value),
            XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(value),
            _ => throw new ArgumentOutOfRangeException(),
        };
    }

    internal void SetCellValue(Point point, XLCellValue cellValue)
    {
        ref readonly XLValueSliceContent original = ref this._values[point];

        double value;
        if (cellValue.Type == XLDataType.Text)
        {
            if (original.Type == XLDataType.Text)
            {
                // Change references. Increase first and then decrease to have fewer shuffles assigning same value to a cell.
                int originalStringId = (int)original.Value;
                value = this._sst.IncreaseRef(cellValue.GetText(), original.Inline);
                this._sst.DecreaseRef(originalStringId);
            }
            else
            {
                // The original value wasn't a text -> just increase ref count to a new text
                value = this._sst.IncreaseRef(cellValue.GetText(), original.Inline);
            }
        }
        else
        {
            // New value isn't a text
            if (original.Type == XLDataType.Text)
            {
                // Dereference original text
                int originalStringId = (int)original.Value;
                this._sst.DecreaseRef(originalStringId);
            }

            if (cellValue.IsUnifiedNumber)
            {
                value = cellValue.GetUnifiedNumber();
            }
            else if (cellValue.IsBoolean)
            {
                value = cellValue.GetBoolean() ? 1 : 0;
            }
            else if (cellValue.IsError)
            {
                value = (int)cellValue.GetError();
            }
            else
            {
                value = 0; // blank
            }
        }

        XLValueSliceContent modified = new(value, cellValue.Type, original.Inline);
        this._values.Set(point, in modified);
    }

    internal XLImmutableRichText? GetRichText(Point point)
    {
        ref readonly XLValueSliceContent cellValue = ref this._values[point];
        if (cellValue.Type != XLDataType.Text)
        {
            return null;
        }

        double value = cellValue.Value;
        return this._sst.GetRichText((int)value);
    }

    internal void SetRichText(Point point, XLImmutableRichText richText)
    {
        ArgumentNullException.ThrowIfNull(richText);

        ref readonly XLValueSliceContent original = ref this._values[point];

        // If original value was a text (no matter if plain or rich text),
        // dereference because it's being replaced.
        if (original.Type == XLDataType.Text)
        {
            int originalId = (int)original.Value;
            this._sst.DecreaseRef(originalId);
        }

        int richTextId = this._sst.IncreaseRef(richText, original.Inline);
        XLValueSliceContent modified = new(richTextId, XLDataType.Text, original.Inline);
        this._values.Set(point, modified);
    }

    internal bool GetShareString(Point point) => !this._values[point].Inline;

    internal void SetShareString(Point point, bool shareString)
    {
        bool inlineString = !shareString;
        ref readonly XLValueSliceContent original = ref this._values[point];
        if (original.Inline == inlineString)
        {
            return;
        }

        double cellValue = original.Value;
        if (original.Type == XLDataType.Text)
        {
            // Because inline is a part of SST, we have to update stringIds when inline flag changes.
            int originalStringId = (int)cellValue;
            XLImmutableRichText? richText = this._sst.GetRichText(originalStringId);
            if (richText is not null)
            {
                // Cell is storing rich text
                this._sst.DecreaseRef(originalStringId);
                cellValue = this._sst.IncreaseRef(richText, inlineString);
            }
            else
            {
                // Cell is storing plain text.
                string originalString = this._sst[originalStringId];
                this._sst.DecreaseRef(originalStringId);
                cellValue = this._sst.IncreaseRef(originalString, inlineString);
            }
        }

        XLValueSliceContent modified = new(cellValue, original.Type, inlineString);
        this._values.Set(point, in modified);
    }

    internal int GetShareStringId(Point point)
    {
        ref readonly XLValueSliceContent value = ref this._values[point];
        if (value.Type != XLDataType.Text)
        {
            throw new InvalidOperationException(
                $"Asking for a shared string id of a non-text cell {point}."
            );
        }

        return (int)this._values[point].Value;
    }

    /// <summary>
    /// Prepare for worksheet removal, dereference all tests in a slice.
    /// </summary>
    internal void DereferenceSlice() => this.DereferenceTextInRange(Area.Full);

    private void DereferenceTextInRange(Area range)
    {
        // Dereference all texts in the range, so the ref count is kept correct.
        using IEnumerator<Point> e = this._values.GetEnumerator(range);
        while (e.MoveNext())
        {
            ref readonly XLValueSliceContent value = ref this._values[e.Current];
            if (value.Type == XLDataType.Text)
            {
                this._sst.DecreaseRef((int)value.Value);
                XLValueSliceContent blank = new(0, XLDataType.Blank, value.Inline);
                this._values.Set(e.Current, in blank);
            }
        }
    }

    private readonly record struct XLValueSliceContent
    {
        /// <summary>
        /// A cell value in a very compact representation. The value is interpreted depending on a type.
        /// </summary>
        internal readonly double Value;

        /// <summary>
        /// Type of a cell <see cref="Value"/>.
        /// </summary>
        internal readonly XLDataType Type;
        internal readonly bool Inline;

        internal XLValueSliceContent(double value, XLDataType type, bool inline)
        {
            this.Value = value;
            this.Type = type;
            this.Inline = inline;
        }
    }
}

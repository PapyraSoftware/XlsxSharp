using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Index;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLRanges : IXLRanges, IEnumerable<XLRange>
{
    private readonly XLWorkbook _workbook;

    /// <summary>
    /// Normally, XLRanges collection includes ranges from a single worksheet, but not necessarily.
    /// </summary>
    private readonly Dictionary<IXLWorksheet, IXLRangeIndex<XLRange>> _indexes;
    private IEnumerable<XLRange> Ranges => this._indexes.Values.SelectMany(index => index.GetAll());

    private IXLRangeIndex<XLRange> GetRangeIndex(IXLWorksheet worksheet)
    {
        if (!this._indexes.TryGetValue(worksheet, out IXLRangeIndex<XLRange> rangeIndex))
        {
            rangeIndex = new XLRangeIndex<XLRange>(worksheet);
            this._indexes.Add(worksheet, rangeIndex);
        }

        return rangeIndex;
    }

    public XLRanges(XLWorksheet worksheet)
        : this(worksheet.Workbook) { }

    public XLRanges(XLWorkbook workbook)
    {
        this._workbook = workbook;
        this._indexes = new Dictionary<IXLWorksheet, IXLRangeIndex<XLRange>>();
    }

    internal XLCellFormat Format
    {
        get
        {
            XLWorksheet? sheet = this.Ranges.FirstOrDefault()?.Worksheet;
            SheetArea[] areas = [.. this.Ranges.Select(x => SheetArea.From(x.RangeAddress))];
            return XLCellFormat.ForAreas(this._workbook, areas, sheet);
        }
    }

    #region IXLRanges Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    IXLCells IXLRanges.Cells() => this.Cells();

    public IXLRanges Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this.Ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Add(XLRange range)
    {
        if (this.GetRangeIndex(range.Worksheet).Add(range))
        {
            this.Count++;
        }
    }

    public void Add(IXLRangeBase range) => this.Add((XLRange)range.AsRange());

    public void Add(IXLCell cell) => this.Add(cell.AsRange());

    public bool Remove(IXLRange range)
    {
        if (this.GetRangeIndex(range.Worksheet).Remove(range.RangeAddress))
        {
            this.Count--;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Removes ranges matching the criteria from the collection, optionally releasing their event handlers.
    /// </summary>
    /// <param name="match">Criteria to filter ranges. Only those ranges that satisfy the criteria will be removed.
    /// Null means the entire collection should be cleared.</param>
    /// <param name="releaseEventHandlers">Specify whether or not should removed ranges be unsubscribed from
    /// row/column shifting events. Until ranges are unsubscribed they cannot be collected by GC.</param>
    public void RemoveAll(Predicate<IXLRange>? match = null, bool releaseEventHandlers = true)
    {
        foreach (IXLRangeIndex<XLRange> index in this._indexes.Values)
        {
            this.Count -= index.RemoveAll(match ?? (_ => true));
        }
    }

    public int Count { get; private set; }

    public IEnumerator<XLRange> GetEnumerator() =>
        this
            .Ranges.OrderBy(r => r.Worksheet.Position)
            .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
            .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
            .GetEnumerator();

    IEnumerator<IXLRange> IEnumerable<IXLRange>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public bool Contains(IXLCell cell) => this.GetIntersectedRanges((XLAddress)cell.Address).Any();

    public bool Contains(IXLRange range) =>
        this.GetIntersectedRanges((XLRangeAddress)range.RangeAddress).Any(r => r.Contains(range));

    /// <summary>
    /// Filter ranges from a collection that intersect the specified address. Is much more efficient
    /// that using Linq expression .Where().
    /// </summary>
    public IEnumerable<IXLRange> GetIntersectedRanges(IXLRangeAddress rangeAddress)
    {
        XLRangeAddress xlRangeAddress = (XLRangeAddress)rangeAddress;
        return this.GetIntersectedRanges(in xlRangeAddress);
    }

    internal IEnumerable<IXLRange> GetIntersectedRanges(in XLRangeAddress rangeAddress) =>
        this.GetRangeIndex(rangeAddress.Worksheet!).GetIntersectedRanges(rangeAddress);

    /// <summary>
    /// Filter ranges from a collection that intersect the specified address. Is much more efficient
    /// that using Linq expression .Where().
    /// </summary>
    public IEnumerable<IXLRange> GetIntersectedRanges(IXLAddress address)
    {
        XLAddress xlAddress = (XLAddress)address;
        return this.GetIntersectedRanges(in xlAddress);
    }

    internal IEnumerable<IXLRange> GetIntersectedRanges(in XLAddress address) =>
        this.GetRangeIndex(address.Worksheet).GetIntersectedRanges(address);

    public IEnumerable<IXLRange> GetIntersectedRanges(IXLCell cell) =>
        this.GetIntersectedRanges(cell.Address);

    public IEnumerable<IXLDataValidation> DataValidation =>
        this.Ranges.Select(range => range.GetDataValidation()).Where(dv => dv != null);

    public IXLRanges AddToNamed(string rangeName) => this.AddToNamed(rangeName, XLScope.Workbook);

    public IXLRanges AddToNamed(string rangeName, XLScope scope) =>
        this.AddToNamed(rangeName, XLScope.Workbook, null);

    public IXLRanges AddToNamed(string rangeName, XLScope scope, string? comment)
    {
        this.Ranges.ForEach(r => r.AddToNamed(rangeName, scope, comment));
        return this;
    }

    public XLCellValue Value
    {
        set => this.Ranges.ForEach(r => r.Value = value);
    }

    public IXLRanges SetValue(XLCellValue value)
    {
        this.Ranges.ForEach(r => r.SetValue(value));
        return this;
    }

    public XLCells Cells()
    {
        XLCells cells = new(this._workbook, false, XLCellsUsedOptions.AllContents);
        foreach (XLRange container in this.Ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(this._workbook, true, XLCellsUsedOptions.AllContents);
        foreach (XLRange container in this.Ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._workbook, true, options);
        foreach (XLRange container in this.Ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    #endregion IXLRanges Members

    public override string ToString()
    {
        string retVal = this.Ranges.Aggregate(string.Empty, (agg, r) => agg + (r.ToString() + ","));
        if (retVal.Length > 0)
        {
            retVal = retVal.Substring(0, retVal.Length - 1);
        }

        return retVal;
    }

    public override bool Equals(object obj) => this.Equals(obj as XLRanges);

    public bool Equals(XLRanges? other)
    {
        if (other == null)
        {
            return false;
        }

        return this.Ranges.Count() == other.Ranges.Count()
            && this.Ranges.Select(thisRange => Enumerable.Contains(other.Ranges, thisRange))
                .All(foundOne => foundOne);
    }

    public override int GetHashCode() =>
        this.Ranges.Aggregate(0, (current, r) => current ^ r.GetHashCode());

    public IXLDataValidation CreateDataValidation()
    {
        XLRange firstRange = this.Ranges.First();
        XLDataValidation dataValidation = firstRange.Worksheet.DataValidations.Create(
            firstRange.SheetRange
        );
        foreach (XLRange range in this.Ranges.Skip(1))
        {
            dataValidation.AddRange(range);
        }

        return dataValidation;
    }

    [Obsolete("Use CreateDataValidation() instead.")]
    public IXLDataValidation SetDataValidation() => this.CreateDataValidation();

    public void Select()
    {
        foreach (XLRange range in this)
        {
            range.Select();
        }
    }

    public IXLRanges Consolidate()
    {
        XLRangeConsolidationEngine engine = new(this._workbook, this);
        return engine.Consolidate();
    }
}

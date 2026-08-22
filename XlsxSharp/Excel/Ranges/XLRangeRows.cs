#nullable disable

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLRangeRows : IXLRangeRows
{
    private readonly XLWorksheet _worksheet;
    private readonly List<XLRangeRow> _ranges = [];

    public XLRangeRows(XLWorksheet worksheet) => this._worksheet = worksheet;

    internal XLCellFormat Format
    {
        get
        {
            SheetArea[] areas = [.. this._ranges.Select(x => SheetArea.From(x.RangeAddress))];
            return XLCellFormat.ForAreas(this._worksheet.Workbook, areas, null);
        }
    }

    #region IXLRangeRows Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    public IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this._ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        this._ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        this._ranges.Clear();
    }

    public void Add(IXLRangeRow range) => this._ranges.Add((XLRangeRow)range);

    public IEnumerator<IXLRangeRow> GetEnumerator() =>
        this
            ._ranges.Cast<IXLRangeRow>()
            .OrderBy(r => r.Worksheet.Position)
            .ThenBy(r => r.RowNumber())
            .GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLCells Cells()
    {
        XLCells cells = new(this._worksheet, false, XLCellsUsedOptions.AllContents);
        foreach (XLRangeRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(this._worksheet, true, XLCellsUsedOptions.AllContents);
        foreach (XLRangeRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._worksheet, true, options);
        foreach (XLRangeRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public void Select()
    {
        foreach (IXLRangeRow range in this)
        {
            range.Select();
        }
    }

    #endregion IXLRangeRows Members
}

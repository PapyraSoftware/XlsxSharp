using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Tables;

internal class XLTableRows : IXLTableRows
{
    private readonly XLWorksheet _worksheet;
    private readonly List<XLTableRow> _ranges = [];

    public XLTableRows(XLWorksheet worksheet) => this._worksheet = worksheet;

    #region IXLTableRows Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this._ranges.ForEach(r => r.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        this._ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        this._ranges.Clear();
    }

    public void Add(IXLTableRow tableRow) => this._ranges.Add((XLTableRow)tableRow);

    public IEnumerator<IXLTableRow> GetEnumerator()
    {
        List<IXLTableRow> retList = [];
        this._ranges.ForEach(retList.Add);
        return retList.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLCells Cells()
    {
        XLCells cells = new(this._worksheet, false, XLCellsUsedOptions.AllContents);
        foreach (XLTableRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(this._worksheet, true, XLCellsUsedOptions.AllContents);
        foreach (XLTableRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(Boolean includeFormats) =>
        this.CellsUsed(includeFormats ? XLCellsUsedOptions.All : XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._worksheet, false, options);
        foreach (XLTableRow container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public void Select()
    {
        foreach (IXLTableRow range in this)
        {
            range.Select();
        }
    }

    #endregion IXLTableRows Members

    internal XLCellFormat Format
    {
        get
        {
            SheetArea[] rowAreas = [.. this._ranges.Select(x => SheetArea.From(x.RangeAddress))];
            return XLCellFormat.ForTableRows(this._worksheet, rowAreas);
        }
    }
}

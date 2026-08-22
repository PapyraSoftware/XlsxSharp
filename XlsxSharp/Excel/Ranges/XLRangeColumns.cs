#nullable disable

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLRangeColumns : IXLRangeColumns
{
    private readonly XLWorksheet _worksheet;
    private readonly List<XLRangeColumn> _ranges = [];

    public XLRangeColumns(XLWorksheet worksheet) => this._worksheet = worksheet;

    internal XLCellFormat Format
    {
        get
        {
            SheetArea[] columns = [.. this._ranges.Select(x => SheetArea.From(x.RangeAddress))];
            return XLCellFormat.ForAreas(this._worksheet.Workbook, columns, null);
        }
    }

    #region IXLRangeColumns Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this._ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        this._ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
        this._ranges.Clear();
    }

    public void Add(IXLRangeColumn range) => this._ranges.Add((XLRangeColumn)range);

    public IEnumerator<IXLRangeColumn> GetEnumerator() =>
        this
            ._ranges.Cast<IXLRangeColumn>()
            .OrderBy(r => r.Worksheet.Position)
            .ThenBy(r => r.ColumnNumber())
            .GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLCells Cells()
    {
        XLCells cells = new(
            this._worksheet,
            usedCellsOnly: false,
            options: XLCellsUsedOptions.AllContents
        );
        foreach (XLRangeColumn container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(
            this._worksheet,
            usedCellsOnly: true,
            options: XLCellsUsedOptions.AllContents
        );
        foreach (XLRangeColumn container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._worksheet, usedCellsOnly: true, options: options);
        foreach (XLRangeColumn container in this._ranges)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public void Select()
    {
        foreach (IXLRangeColumn range in this)
        {
            range.Select();
        }
    }

    #endregion IXLRangeColumns Members
}

#nullable disable

using System;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel.Tables;

internal class XLTableRow : XLRangeRow, IXLTableRow
{
    private readonly XLTableRange _tableRange;

    public XLTableRow(XLTableRange tableRange, XLRangeRow rangeRow)
        : base(rangeRow.RangeAddress)
    {
        this._tableRange = tableRange;
    }

    #region IXLTableRow Members

    public IXLCell Field(Int32 index)
    {
        return this.Cell(index + 1);
    }

    public IXLCell Field(String name)
    {
        Int32 fieldIndex = this._tableRange.Table.GetFieldIndex(name);
        return this.Cell(fieldIndex + 1);
    }

    public new IXLTableRow Sort()
    {
        return this.SortLeftToRight();
    }

    public new IXLTableRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    )
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    #endregion IXLTableRow Members

    private XLTableRow RowShift(Int32 rowsToShift)
    {
        return this._tableRange.Row(
            this.RowNumber() - this._tableRange.FirstRow().RowNumber() + 1 + rowsToShift
        );
    }

    #region XLTableRow Above

    IXLTableRow IXLTableRow.RowAbove()
    {
        return this.RowAbove();
    }

    IXLTableRow IXLTableRow.RowAbove(Int32 step)
    {
        return this.RowAbove(step);
    }

    public new XLTableRow RowAbove()
    {
        return this.RowAbove(1);
    }

    public new XLTableRow RowAbove(Int32 step)
    {
        return this.RowShift(step * -1);
    }

    #endregion XLTableRow Above

    #region XLTableRow Below

    IXLTableRow IXLTableRow.RowBelow()
    {
        return this.RowBelow();
    }

    IXLTableRow IXLTableRow.RowBelow(Int32 step)
    {
        return this.RowBelow(step);
    }

    public new XLTableRow RowBelow()
    {
        return this.RowBelow(1);
    }

    public new XLTableRow RowBelow(Int32 step)
    {
        return this.RowShift(step);
    }

    #endregion XLTableRow Below

    public new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public new IXLTableRows InsertRowsAbove(int numberOfRows)
    {
        return XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsAbove,
            this._tableRange,
            numberOfRows,
            !this._tableRange.Table.ShowTotalsRow
        );
    }

    public new IXLTableRows InsertRowsBelow(int numberOfRows)
    {
        return XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsBelow,
            this._tableRange,
            numberOfRows,
            !this._tableRange.Table.ShowTotalsRow
        );
    }

    public new void Delete()
    {
        this.Delete(XLShiftDeletedCells.ShiftCellsUp);
    }
}

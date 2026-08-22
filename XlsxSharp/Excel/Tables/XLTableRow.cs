#nullable disable

using System;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel.Tables;

internal class XLTableRow : XLRangeRow, IXLTableRow
{
    private readonly XLTableRange _tableRange;

    public XLTableRow(XLTableRange tableRange, XLRangeRow rangeRow)
        : base(rangeRow.RangeAddress) => this._tableRange = tableRange;

    #region IXLTableRow Members

    public IXLCell Field(int index) => this.Cell(index + 1);

    public IXLCell Field(string name)
    {
        int fieldIndex = this._tableRange.Table.GetFieldIndex(name);
        return this.Cell(fieldIndex + 1);
    }

    public new IXLTableRow Sort() => this.SortLeftToRight();

    public new IXLTableRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    #endregion IXLTableRow Members

    private XLTableRow RowShift(int rowsToShift) =>
        this._tableRange.Row(
            this.RowNumber() - this._tableRange.FirstRow().RowNumber() + 1 + rowsToShift
        );

    #region XLTableRow Above

    IXLTableRow IXLTableRow.RowAbove() => this.RowAbove();

    IXLTableRow IXLTableRow.RowAbove(int step) => this.RowAbove(step);

    public new XLTableRow RowAbove() => this.RowAbove(1);

    public new XLTableRow RowAbove(int step) => this.RowShift(step * -1);

    #endregion XLTableRow Above

    #region XLTableRow Below

    IXLTableRow IXLTableRow.RowBelow() => this.RowBelow();

    IXLTableRow IXLTableRow.RowBelow(int step) => this.RowBelow(step);

    public new XLTableRow RowBelow() => this.RowBelow(1);

    public new XLTableRow RowBelow(int step) => this.RowShift(step);

    #endregion XLTableRow Below

    public new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public new IXLTableRows InsertRowsAbove(int numberOfRows) =>
        XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsAbove,
            this._tableRange,
            numberOfRows,
            !this._tableRange.Table.ShowTotalsRow
        );

    public new IXLTableRows InsertRowsBelow(int numberOfRows) =>
        XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsBelow,
            this._tableRange,
            numberOfRows,
            !this._tableRange.Table.ShowTotalsRow
        );

    public new void Delete() => this.Delete(XLShiftDeletedCells.ShiftCellsUp);
}

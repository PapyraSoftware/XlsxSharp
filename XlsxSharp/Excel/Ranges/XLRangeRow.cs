using System;
using System.Linq;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel;

internal class XLRangeRow : XLRangeBase, IXLRangeRow
{
    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLRangeRow(XLRangeAddress rangeAddress)
        : base(rangeAddress) { }

    internal override XLCellFormat Format =>
        XLCellFormat.ForRange(this.Worksheet, this.RangeAddress);

    #region IXLRangeRow Members

    IXLCells IXLRangeRow.Cells(string cellsInRow) => this.Cells(cellsInRow);

    public IXLCell Cell(int column) => this.Cell(1, column);

    public override XLCell Cell(string columnLetter) => this.Cell(1, columnLetter);

    IXLCell IXLRangeRow.Cell(string columnLetter) => this.Cell(columnLetter);

    public void Delete() => this.Delete(XLShiftDeletedCells.ShiftCellsUp);

    public IXLCells InsertCellsAfter(int numberOfColumns) =>
        this.InsertCellsAfter(numberOfColumns, true);

    public IXLCells InsertCellsAfter(int numberOfColumns, bool expandRange) =>
        this.InsertColumnsAfter(numberOfColumns, expandRange).Cells();

    public IXLCells InsertCellsBefore(int numberOfColumns) =>
        this.InsertCellsBefore(numberOfColumns, false);

    public IXLCells InsertCellsBefore(int numberOfColumns, bool expandRange) =>
        this.InsertColumnsBefore(numberOfColumns, expandRange).Cells();

    public override XLCells Cells(string cellsInRow)
    {
        XLCells retVal = new(this.Worksheet, false, XLCellsUsedOptions.AllContents);
        string[] rangePairs = cellsInRow.Split(',');
        foreach (string pair in rangePairs)
        {
            retVal.Add(this.Range(pair.Trim()).RangeAddress);
        }

        return retVal;
    }

    public IXLCells Cells(int firstColumn, int lastColumn) =>
        this.Cells(firstColumn + ":" + lastColumn);

    public IXLCells Cells(string firstColumn, string lastColumn) =>
        this.Cells(
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(firstColumn)
                + ":"
                + XlsxSharp.XLHelper.GetColumnNumberFromLetter(lastColumn)
        );

    public int CellCount() =>
        this.RangeAddress.LastAddress.ColumnNumber
        - this.RangeAddress.FirstAddress.ColumnNumber
        + 1;

    public new IXLRangeRow Sort() => this.SortLeftToRight();

    public new IXLRangeRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    )
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    public IXLRangeRow CopyTo(IXLCell target)
    {
        base.CopyTo((XLCell)target);

        int lastRowNumber = target.Address.RowNumber + this.RowCount() - 1;
        if (lastRowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRowNumber = XlsxSharp.XLHelper.MaxRowNumber;
        }

        int lastColumnNumber = target.Address.ColumnNumber + this.ColumnCount() - 1;
        if (lastColumnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumnNumber = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        return target
            .Worksheet.Range(
                target.Address.RowNumber,
                target.Address.ColumnNumber,
                lastRowNumber,
                lastColumnNumber
            )
            .Row(1);
    }

    public new IXLRangeRow CopyTo(IXLRangeBase target)
    {
        base.CopyTo(target);
        int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + this.RowCount() - 1;
        if (lastRowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRowNumber = XlsxSharp.XLHelper.MaxRowNumber;
        }

        int lastColumnNumber =
            target.RangeAddress.LastAddress.ColumnNumber + this.ColumnCount() - 1;
        if (lastColumnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumnNumber = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        return target
            .Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.LastAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber
            )
            .Row(1);
    }

    public IXLRangeRow Row(int start, int end) => this.Range(1, start, 1, end).Row(1);

    public IXLRangeRow Row(IXLCell start, IXLCell end) =>
        this.Row(start.Address.ColumnNumber, end.Address.ColumnNumber);

    public IXLRangeRows Rows(string rows)
    {
        XLRangeRows retVal = new(this.Worksheet);
        string[] columnPairs = rows.Split(',');
        foreach (string trimmedPair in columnPairs.Select(pair => pair.Trim()))
        {
            string firstColumn;
            string lastColumn;
            if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
            {
                string[] columnRange = trimmedPair.Contains('-')
                    ? trimmedPair.Replace('-', ':').Split(':')
                    : trimmedPair.Split(':');
                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = trimmedPair;
                lastColumn = trimmedPair;
            }

            retVal.Add(this.Range(firstColumn, lastColumn).FirstRow());
        }

        return retVal;
    }

    public IXLRow WorksheetRow() => this.Worksheet.Row(this.RangeAddress.FirstAddress.RowNumber);

    #endregion IXLRangeRow Members
    public override XLRangeType RangeType => XLRangeType.RangeRow;

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted) =>
        this.RangeAddress = (XLRangeAddress)
            this.ShiftColumns(this.RangeAddress, range, columnsShifted);

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted) =>
        this.RangeAddress = (XLRangeAddress)this.ShiftRows(this.RangeAddress, range, rowsShifted);

    public IXLRange Range(int firstColumn, int lastColumn) =>
        this.Range(1, firstColumn, 1, lastColumn);

    public override XLRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
            {
                rangeAddressStr = rangeAddressStr.Replace('-', ':');
            }

            string[] arrRange = rangeAddressStr.Split(':');
            string firstPart = arrRange[0];
            string secondPart = arrRange[1];
            rangeAddressToUse =
                this.FixRowAddress(firstPart) + ":" + this.FixRowAddress(secondPart);
        }
        else
        {
            rangeAddressToUse = this.FixRowAddress(rangeAddressStr);
        }

        XLRangeAddress rangeAddress = new(this.Worksheet, rangeAddressToUse);
        return this.Range(rangeAddress);
    }

    public int CompareTo(XLRangeRow otherRow, IXLSortElements columnsToSort)
    {
        foreach (IXLSortElement e in columnsToSort)
        {
            XLCell thisCell = (XLCell)this.Cell(e.ElementNumber);
            XLCell otherCell = (XLCell)otherRow.Cell(e.ElementNumber);
            int comparison;
            bool thisCellIsBlank = thisCell.IsEmpty();
            bool otherCellIsBlank = otherCell.IsEmpty();
            if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
            {
                if (thisCellIsBlank && otherCellIsBlank)
                {
                    comparison = 0;
                }
                else
                {
                    if (thisCellIsBlank)
                    {
                        comparison = e.SortOrder == XLSortOrder.Ascending ? 1 : -1;
                    }
                    else
                    {
                        comparison = e.SortOrder == XLSortOrder.Ascending ? -1 : 1;
                    }
                }
            }
            else
            {
                if (thisCell.DataType == otherCell.DataType)
                {
                    switch (thisCell.DataType)
                    {
                        case XLDataType.Text:
                            comparison = e.MatchCase
                                ? thisCell.GetText().CompareTo(otherCell.GetText())
                                : String.Compare(thisCell.GetText(), otherCell.GetText(), true);
                            break;

                        case XLDataType.TimeSpan:
                            comparison = thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
                            break;

                        case XLDataType.DateTime:
                            comparison = thisCell.GetDateTime().CompareTo(otherCell.GetDateTime());
                            break;

                        case XLDataType.Number:
                            comparison = thisCell.GetDouble().CompareTo(otherCell.GetDouble());
                            break;

                        case XLDataType.Boolean:
                            comparison = thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
                            break;

                        default:
                            throw new NotImplementedException();
                    }
                }
                else if (thisCell.Value.IsUnifiedNumber && otherCell.Value.IsUnifiedNumber)
                {
                    comparison = thisCell
                        .Value.GetUnifiedNumber()
                        .CompareTo(otherCell.Value.GetUnifiedNumber());
                }
                else if (e.MatchCase)
                {
                    comparison = String.Compare(thisCell.GetString(), otherCell.GetString(), true);
                }
                else
                {
                    comparison = thisCell.GetString().CompareTo(otherCell.GetString());
                }
            }

            if (comparison != 0)
            {
                return e.SortOrder == XLSortOrder.Ascending ? comparison : -comparison;
            }
        }

        return 0;
    }

    private XLRangeRow RowShift(Int32 rowsToShift)
    {
        Int32 rowNum = this.RowNumber() + rowsToShift;

        XLRange range = this.Worksheet.Range(
            rowNum,
            this.RangeAddress.FirstAddress.ColumnNumber,
            rowNum,
            this.RangeAddress.LastAddress.ColumnNumber
        );

        return range.FirstRow();
    }

    #region XLRangeRow Above

    IXLRangeRow IXLRangeRow.RowAbove() => this.RowAbove();

    IXLRangeRow IXLRangeRow.RowAbove(Int32 step) => this.RowAbove(step);

    public XLRangeRow RowAbove() => this.RowAbove(1);

    public XLRangeRow RowAbove(Int32 step) => this.RowShift(step * -1);

    #endregion XLRangeRow Above

    #region XLRangeRow Below

    IXLRangeRow IXLRangeRow.RowBelow() => this.RowBelow();

    IXLRangeRow IXLRangeRow.RowBelow(Int32 step) => this.RowBelow(step);

    public XLRangeRow RowBelow() => this.RowBelow(1);

    public XLRangeRow RowBelow(Int32 step) => this.RowShift(step);

    #endregion XLRangeRow Below

    public new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents) =>
        this.Row(
            (this as IXLRangeBase).FirstCellUsed(options),
            (this as IXLRangeBase).LastCellUsed(options)
        );
}

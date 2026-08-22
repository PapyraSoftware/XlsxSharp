#nullable disable

using System;
using System.Linq;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

internal class XLRangeColumn : XLRangeBase, IXLRangeColumn
{
    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLRangeColumn(XLRangeAddress rangeAddress)
        : base(rangeAddress) { }

    internal override XLCellFormat Format =>
        XLCellFormat.ForRange(this.Worksheet, this.RangeAddress);

    #region IXLRangeColumn Members

    IXLCell IXLRangeColumn.Cell(int rowNumber)
    {
        return this.Cell(rowNumber);
    }

    IXLCells IXLRangeColumn.Cells(string cellsInColumn) => this.Cells(cellsInColumn);

    public override XLCells Cells(string cellsInColumn)
    {
        XLCells retVal = new(this.Worksheet, false, XLCellsUsedOptions.AllContents);
        string[] rangePairs = cellsInColumn.Split(',');
        foreach (string pair in rangePairs)
        {
            retVal.Add(this.Range(pair.Trim()).RangeAddress);
        }

        return retVal;
    }

    public IXLCells Cells(int firstRow, int lastRow)
    {
        return this.Cells(firstRow + ":" + lastRow);
    }

    public void Delete()
    {
        this.Delete(true);
    }

    internal void Delete(Boolean deleteTableField)
    {
        if (deleteTableField && this.IsTableColumn())
        {
            XLTable table = this.Table as XLTable;
            if (!this.Cell(1).Value.TryGetText(out string firstCellValue))
            {
                throw new InvalidOperationException("Top cell doesn't contain a text.");
            }

            if (!table.FieldNames.ContainsKey(firstCellValue))
            {
                throw new InvalidOperationException($"Field {firstCellValue} not found.");
            }

            XLTableField field = table
                .Fields.Cast<XLTableField>()
                .Single(f => f.Name == firstCellValue);
            field.Delete(false);
        }

        this.Delete(XLShiftDeletedCells.ShiftCellsLeft);
    }

    public IXLCells InsertCellsAbove(int numberOfRows)
    {
        return this.InsertCellsAbove(numberOfRows, false);
    }

    public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange)
    {
        return this.InsertRowsAbove(numberOfRows, expandRange).Cells();
    }

    public IXLCells InsertCellsBelow(int numberOfRows)
    {
        return this.InsertCellsBelow(numberOfRows, true);
    }

    public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange)
    {
        return this.InsertRowsBelow(numberOfRows, expandRange).Cells();
    }

    public int CellCount()
    {
        return this.RangeAddress.LastAddress.RowNumber
            - this.RangeAddress.FirstAddress.RowNumber
            + 1;
    }

    public IXLRangeColumn Sort(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    )
    {
        base.Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    public IXLRangeColumn CopyTo(IXLCell target)
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
            .Column(1);
    }

    public new IXLRangeColumn CopyTo(IXLRangeBase target)
    {
        base.CopyTo(target);

        int lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + this.RowCount() - 1;
        if (lastRowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRowNumber = XlsxSharp.XLHelper.MaxRowNumber;
        }

        int lastColumnNumber =
            target.RangeAddress.FirstAddress.ColumnNumber + this.ColumnCount() - 1;
        if (lastColumnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumnNumber = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        return target
            .Worksheet.Range(
                target.RangeAddress.FirstAddress.RowNumber,
                target.RangeAddress.FirstAddress.ColumnNumber,
                lastRowNumber,
                lastColumnNumber
            )
            .Column(1);
    }

    public IXLRangeColumn Column(int start, int end)
    {
        return this.Range(start, end).FirstColumn();
    }

    public IXLRangeColumn Column(IXLCell start, IXLCell end)
    {
        return this.Column(start.Address.RowNumber, end.Address.RowNumber);
    }

    public IXLRangeColumns Columns(string columns)
    {
        XLRangeColumns retVal = new(this.Worksheet);
        string[] rowPairs = columns.Split(',');
        foreach (string trimmedPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
            {
                string[] rowRange = trimmedPair.Split(':', '-');

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = trimmedPair;
                lastRow = trimmedPair;
            }

            retVal.Add(this.Range(firstRow, lastRow).FirstColumn());
        }

        return retVal;
    }

    public IXLColumn WorksheetColumn()
    {
        return this.Worksheet.Column(this.RangeAddress.FirstAddress.ColumnNumber);
    }

    #endregion IXLRangeColumn Members

    public override XLRangeType RangeType
    {
        get { return XLRangeType.RangeColumn; }
    }

    public XLCell Cell(int row)
    {
        return this.Cell(row, 1);
    }

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        this.RangeAddress = (XLRangeAddress)
            this.ShiftColumns(this.RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        this.RangeAddress = (XLRangeAddress)this.ShiftRows(this.RangeAddress, range, rowsShifted);
    }

    public XLRange Range(int firstRow, int lastRow)
    {
        return this.Range(firstRow, 1, lastRow, 1);
    }

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
                this.FixColumnAddress(firstPart) + ":" + this.FixColumnAddress(secondPart);
        }
        else
        {
            rangeAddressToUse = this.FixColumnAddress(rangeAddressStr);
        }

        XLRangeAddress rangeAddress = new(this.Worksheet, rangeAddressToUse);
        return this.Range(rangeAddress);
    }

    public int CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
    {
        foreach (IXLSortElement e in rowsToSort)
        {
            XLCell thisCell = this.Cell(e.ElementNumber);
            XLCell otherCell = otherColumn.Cell(e.ElementNumber);
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
                    if (thisCell.DataType == XLDataType.Blank)
                    {
                        comparison = 0;
                    }
                    else if (thisCell.DataType == XLDataType.Boolean)
                    {
                        comparison = thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
                    }
                    else if (thisCell.DataType == XLDataType.Text)
                    {
                        comparison = e.MatchCase
                            ? thisCell.GetText().CompareTo(otherCell.GetText())
                            : String.Compare(thisCell.GetText(), otherCell.GetText(), true);
                    }
                    else if (thisCell.DataType == XLDataType.Error)
                    {
                        comparison = 0; // Errors are incomparable
                    }
                    else
                    {
                        comparison = thisCell
                            .CachedValue.GetUnifiedNumber()
                            .CompareTo(thisCell.CachedValue.GetUnifiedNumber());
                    }
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
                return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
            }
        }

        return 0;
    }

    private XLRangeColumn ColumnShift(Int32 columnsToShift)
    {
        Int32 columnNumber = this.ColumnNumber() + columnsToShift;
        return this
            .Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                columnNumber,
                this.RangeAddress.LastAddress.RowNumber,
                columnNumber
            )
            .FirstColumn();
    }

    #region XLRangeColumn Left

    IXLRangeColumn IXLRangeColumn.ColumnLeft()
    {
        return this.ColumnLeft();
    }

    IXLRangeColumn IXLRangeColumn.ColumnLeft(Int32 step)
    {
        return this.ColumnLeft(step);
    }

    public XLRangeColumn ColumnLeft()
    {
        return this.ColumnLeft(1);
    }

    public XLRangeColumn ColumnLeft(Int32 step)
    {
        return this.ColumnShift(step * -1);
    }

    #endregion XLRangeColumn Left

    #region XLRangeColumn Right

    IXLRangeColumn IXLRangeColumn.ColumnRight()
    {
        return this.ColumnRight();
    }

    IXLRangeColumn IXLRangeColumn.ColumnRight(Int32 step)
    {
        return this.ColumnRight(step);
    }

    public XLRangeColumn ColumnRight()
    {
        return this.ColumnRight(1);
    }

    public XLRangeColumn ColumnRight(Int32 step)
    {
        return this.ColumnShift(step);
    }

    #endregion XLRangeColumn Right

    public IXLTable AsTable()
    {
        if (this.IsTableColumn())
        {
            throw new InvalidOperationException("This column is already part of a table.");
        }

        return this.AsRange().AsTable();
    }

    public IXLTable AsTable(string name)
    {
        if (this.IsTableColumn())
        {
            throw new InvalidOperationException("This column is already part of a table.");
        }

        return this.AsRange().AsTable(name);
    }

    public IXLTable CreateTable()
    {
        if (this.IsTableColumn())
        {
            throw new InvalidOperationException("This column is already part of a table.");
        }

        return this.AsRange().CreateTable();
    }

    public IXLTable CreateTable(string name)
    {
        if (this.IsTableColumn())
        {
            throw new InvalidOperationException("This column is already part of a table.");
        }

        return this.AsRange().CreateTable(name);
    }

    public new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
    {
        return this.Column(
            (this as IXLRangeBase).FirstCellUsed(options),
            (this as IXLRangeBase).LastCellUsed(options)
        );
    }

    internal IXLTable Table { get; set; }

    public Boolean IsTableColumn()
    {
        return this.Table != null;
    }
}

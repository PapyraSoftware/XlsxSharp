#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLRange : XLRangeBase, IXLRange
{
    public XLRange(XLRangeAddress rangeAddress, IXLStyle defaultStyle)
        : base(rangeAddress) { }

    internal override XLCellFormat Format =>
        XLCellFormat.ForRange(this.Worksheet, this.RangeAddress);

    public override XLRangeType RangeType => XLRangeType.Range;

    #region IXLRange Members

    IXLRangeRow IXLRange.Row(Int32 row) => this.Row(row);

    IXLRangeColumn IXLRange.Column(Int32 columnNumber) => this.Column(columnNumber);

    IXLRangeColumn IXLRange.Column(String columnLetter) => this.Column(columnLetter);

    public virtual IXLRangeColumns Columns(Func<IXLRangeColumn, Boolean> predicate = null)
    {
        XLRangeColumns retVal = new(this.Worksheet);
        Int32 columnCount = this.ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            XLRangeColumn column = this.Column(c);
            if (predicate == null || predicate(column))
            {
                retVal.Add(column);
            }
        }
        return retVal;
    }

    public virtual IXLRangeColumns Columns(Int32 firstColumn, Int32 lastColumn)
    {
        XLRangeColumns retVal = new(this.Worksheet);

        for (int co = firstColumn; co <= lastColumn; co++)
        {
            retVal.Add(this.Column(co));
        }

        return retVal;
    }

    public virtual IXLRangeColumns Columns(String firstColumn, String lastColumn) =>
        this.Columns(
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(firstColumn),
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(lastColumn)
        );

    public virtual IXLRangeColumns Columns(String columns)
    {
        XLRangeColumns retVal = new(this.Worksheet);
        string[] columnPairs = columns.Split(',');
        foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
        {
            String firstColumn;
            String lastColumn;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] columnRange = XlsxSharp.XLHelper.SplitRange(tPair);

                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = tPair;
                lastColumn = tPair;
            }

            if (Int32.TryParse(firstColumn, out Int32 tmp))
            {
                foreach (
                    IXLRangeColumn col in this.Columns(
                        Int32.Parse(firstColumn),
                        Int32.Parse(lastColumn)
                    )
                )
                {
                    retVal.Add(col);
                }
            }
            else
            {
                foreach (IXLRangeColumn col in this.Columns(firstColumn, lastColumn))
                {
                    retVal.Add(col);
                }
            }
        }
        return retVal;
    }

    IXLCell IXLRange.Cell(int row, int column) => this.Cell(row, column);

    IXLCell IXLRange.Cell(string cellAddressInRange) => this.Cell(cellAddressInRange);

    IXLCell IXLRange.Cell(int row, string column) => this.Cell(row, column);

    IXLCell IXLRange.Cell(IXLAddress cellAddressInRange) => this.Cell(cellAddressInRange);

    IXLRange IXLRange.Range(IXLRangeAddress rangeAddress) => this.Range(rangeAddress);

    IXLRange IXLRange.Range(string rangeAddress) => this.Range(rangeAddress);

    IXLRange IXLRange.Range(IXLCell firstCell, IXLCell lastCell) => this.Range(firstCell, lastCell);

    IXLRange IXLRange.Range(string firstCellAddress, string lastCellAddress) =>
        this.Range(firstCellAddress, lastCellAddress);

    IXLRange IXLRange.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress) =>
        this.Range(firstCellAddress, lastCellAddress);

    IXLRange IXLRange.Range(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    ) => this.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);

    IXLRanges IXLRange.Ranges(string ranges) => this.Ranges(ranges);

    public IXLRangeRows Rows(Func<IXLRangeRow, Boolean> predicate = null)
    {
        XLRangeRows retVal = new(this.Worksheet);
        Int32 rowCount = this.RowCount();
        for (Int32 r = 1; r <= rowCount; r++)
        {
            XLRangeRow row = this.Row(r);
            if (predicate == null || predicate(row))
            {
                retVal.Add(this.Row(r));
            }
        }
        return retVal;
    }

    public IXLRangeRows Rows(Int32 firstRow, Int32 lastRow)
    {
        XLRangeRows retVal = new(this.Worksheet);

        for (int ro = firstRow; ro <= lastRow; ro++)
        {
            retVal.Add(this.Row(ro));
        }

        return retVal;
    }

    public IXLRangeRows Rows(String rows)
    {
        XLRangeRows retVal = new(this.Worksheet);
        string[] rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            String firstRow;
            String lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] rowRange = XlsxSharp.XLHelper.SplitRange(tPair);

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }
            foreach (IXLRangeRow row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
            {
                retVal.Add(row);
            }
        }
        return retVal;
    }

    public void Transpose(XLTransposeOptions transposeOption)
    {
        int rowCount = this.RowCount();
        int columnCount = this.ColumnCount();
        int squareSide = rowCount > columnCount ? rowCount : columnCount;

        XLCell firstCell = this.FirstCell();

        this.MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
        this.TransposeMerged(squareSide);
        this.TransposeRange(squareSide);
        this.RangeAddress = new XLRangeAddress(
            this.RangeAddress.FirstAddress,
            new XLAddress(
                this.Worksheet,
                firstCell.Address.RowNumber + columnCount - 1,
                firstCell.Address.ColumnNumber + rowCount - 1,
                this.RangeAddress.LastAddress.FixedRow,
                this.RangeAddress.LastAddress.FixedColumn
            )
        );

        if (rowCount > columnCount)
        {
            XLRange rng = this.Worksheet.Range(
                this.RangeAddress.LastAddress.RowNumber + 1,
                this.RangeAddress.FirstAddress.ColumnNumber,
                this.RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                this.RangeAddress.LastAddress.ColumnNumber
            );
            rng.Delete(XLShiftDeletedCells.ShiftCellsUp);
        }
        else if (columnCount > rowCount)
        {
            XLRange rng = this.Worksheet.Range(
                this.RangeAddress.FirstAddress.RowNumber,
                this.RangeAddress.LastAddress.ColumnNumber + 1,
                this.RangeAddress.LastAddress.RowNumber,
                this.RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount)
            );
            rng.Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        foreach (IXLCell c in this.Range(1, 1, columnCount, rowCount).Cells())
        {
            XLBorderFormatValue border = ((XLCell)c).GetFormat().Border;
            c.Style.Border.TopBorder = border.Left.Style;
            c.Style.Border.TopBorderColor = border.Left.Color;
            c.Style.Border.LeftBorder = border.Top.Style;
            c.Style.Border.LeftBorderColor = border.Top.Color;
            c.Style.Border.RightBorder = border.Bottom.Style;
            c.Style.Border.RightBorderColor = border.Bottom.Color;
            c.Style.Border.BottomBorder = border.Right.Style;
            c.Style.Border.BottomBorderColor = border.Right.Color;
        }
    }

    public IXLTable AsTable() => this.Worksheet.Table(this, false);

    public IXLTable AsTable(String name) => this.Worksheet.Table(this, name, false);

    IXLTable IXLRange.CreateTable() => this.CreateTable();

    public XLTable CreateTable() => (XLTable)this.Worksheet.Table(this, true, true);

    IXLTable IXLRange.CreateTable(String name) => this.CreateTable(name);

    public XLTable CreateTable(String name) =>
        (XLTable)this.Worksheet.Table(this, name, true, true);

    public IXLTable CreateTable(String name, Boolean setAutofilter) =>
        this.Worksheet.Table(this, name, true, setAutofilter);

    public IXLRange CopyTo(IXLCell target)
    {
        base.CopyTo((XLCell)target);

        Int32 lastRowNumber = target.Address.RowNumber + this.RowCount() - 1;
        if (lastRowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRowNumber = XlsxSharp.XLHelper.MaxRowNumber;
        }

        Int32 lastColumnNumber = target.Address.ColumnNumber + this.ColumnCount() - 1;
        if (lastColumnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumnNumber = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        return target.Worksheet.Range(
            target.Address.RowNumber,
            target.Address.ColumnNumber,
            lastRowNumber,
            lastColumnNumber
        );
    }

    public new IXLRange CopyTo(IXLRangeBase target)
    {
        base.CopyTo(target);

        Int32 lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + this.RowCount() - 1;
        if (lastRowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRowNumber = XlsxSharp.XLHelper.MaxRowNumber;
        }

        Int32 lastColumnNumber =
            target.RangeAddress.FirstAddress.ColumnNumber + this.ColumnCount() - 1;
        if (lastColumnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumnNumber = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        return target.Worksheet.Range(
            target.RangeAddress.FirstAddress.RowNumber,
            target.RangeAddress.FirstAddress.ColumnNumber,
            lastRowNumber,
            lastColumnNumber
        );
    }

    public new IXLRange Sort() => base.Sort().AsRange();

    public new IXLRange Sort(
        String columnsToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    ) => base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();

    public new IXLRange Sort(
        Int32 columnToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    ) => base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();

    public new IXLRange SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        Boolean matchCase = false,
        Boolean ignoreBlanks = true
    ) => base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks).AsRange();

    #endregion IXLRange Members

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted) =>
        this.RangeAddress = (XLRangeAddress)
            this.ShiftColumns(this.RangeAddress, range, columnsShifted);

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted) =>
        this.RangeAddress = (XLRangeAddress)this.ShiftRows(this.RangeAddress, range, rowsShifted);

    IXLRangeColumn IXLRange.FirstColumn(Func<IXLRangeColumn, Boolean> predicate) =>
        this.FirstColumn(predicate);

    internal XLRangeColumn FirstColumn(Func<IXLRangeColumn, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            return this.Column(1);
        }

        Int32 columnCount = this.ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            XLRangeColumn column = this.Column(c);
            if (predicate(column))
            {
                return column;
            }
        }

        return null;
    }

    IXLRangeColumn IXLRange.LastColumn(Func<IXLRangeColumn, Boolean> predicate) =>
        this.LastColumn(predicate);

    internal XLRangeColumn LastColumn(Func<IXLRangeColumn, Boolean> predicate = null)
    {
        Int32 columnCount = this.ColumnCount();
        if (predicate == null)
        {
            return this.Column(columnCount);
        }

        for (Int32 c = columnCount; c >= 1; c--)
        {
            XLRangeColumn column = this.Column(c);
            if (predicate(column))
            {
                return column;
            }
        }

        return null;
    }

    IXLRangeColumn IXLRange.FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate) =>
        this.FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLRangeColumn FirstColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null) =>
        this.FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLRangeColumn IXLRange.FirstColumnUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate
    ) => this.FirstColumnUsed(options, predicate);

    internal XLRangeColumn FirstColumnUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            Int32 firstColumnUsed = this.Worksheet.Internals.CellsCollection.FirstColumnUsed(
                Area.FromRangeAddress(this.RangeAddress),
                options
            );

            return firstColumnUsed == 0
                ? null
                : this.Column(firstColumnUsed - this.RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        Int32 columnCount = this.ColumnCount();
        for (Int32 co = 1; co <= columnCount; co++)
        {
            XLRangeColumn column = this.Column(co);

            if (!column.IsEmpty(options) && predicate(column))
            {
                return column;
            }
        }
        return null;
    }

    IXLRangeColumn IXLRange.LastColumnUsed(Func<IXLRangeColumn, Boolean> predicate) =>
        this.LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLRangeColumn LastColumnUsed(Func<IXLRangeColumn, Boolean> predicate = null) =>
        this.LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLRangeColumn IXLRange.LastColumnUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate
    ) => this.LastColumnUsed(options, predicate);

    internal XLRangeColumn LastColumnUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            Int32 lastColumnUsed = this.Worksheet.Internals.CellsCollection.LastColumnUsed(
                Area.FromRangeAddress(this.RangeAddress),
                options
            );

            return lastColumnUsed == 0
                ? null
                : this.Column(lastColumnUsed - this.RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        Int32 columnCount = this.ColumnCount();
        for (Int32 co = columnCount; co >= 1; co--)
        {
            XLRangeColumn column = this.Column(co);

            if (!column.IsEmpty(options) && predicate(column))
            {
                return column;
            }
        }
        return null;
    }

    IXLRangeRow IXLRange.FirstRow(Func<IXLRangeRow, Boolean> predicate) => this.FirstRow(predicate);

    public XLRangeRow FirstRow(Func<IXLRangeRow, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            return this.Row(1);
        }

        Int32 rowCount = this.RowCount();
        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLRangeRow row = this.Row(ro);
            if (predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLRangeRow IXLRange.LastRow(Func<IXLRangeRow, Boolean> predicate) => this.LastRow(predicate);

    public XLRangeRow LastRow(Func<IXLRangeRow, Boolean> predicate = null)
    {
        Int32 rowCount = this.RowCount();
        if (predicate == null)
        {
            return this.Row(rowCount);
        }

        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            XLRangeRow row = this.Row(ro);
            if (predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLRangeRow IXLRange.FirstRowUsed(Func<IXLRangeRow, Boolean> predicate) =>
        this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLRangeRow FirstRowUsed(Func<IXLRangeRow, Boolean> predicate = null) =>
        this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLRangeRow IXLRange.FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate
    ) => this.FirstRowUsed(options, predicate);

    internal XLRangeRow FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            Int32 rowFromCells = this.Worksheet.Internals.CellsCollection.FirstRowUsed(
                Area.FromRangeAddress(this.RangeAddress),
                options
            );

            return rowFromCells == 0
                ? null
                : this.Row(rowFromCells - this.RangeAddress.FirstAddress.RowNumber + 1);
        }

        Int32 rowCount = this.RowCount();
        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLRangeRow row = this.Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }
        return null;
    }

    IXLRangeRow IXLRange.LastRowUsed(Func<IXLRangeRow, Boolean> predicate) =>
        this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLRangeRow LastRowUsed(Func<IXLRangeRow, Boolean> predicate = null) =>
        this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLRangeRow IXLRange.LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate
    ) => this.LastRowUsed(options, predicate);

    internal XLRangeRow LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            Int32 lastRowUsed = this.Worksheet.Internals.CellsCollection.LastRowUsed(
                Area.FromRangeAddress(this.RangeAddress),
                options
            );

            return lastRowUsed == 0
                ? null
                : this.Row(lastRowUsed - this.RangeAddress.FirstAddress.RowNumber + 1);
        }

        Int32 rowCount = this.RowCount();
        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            XLRangeRow row = this.Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }
        return null;
    }

    IXLRangeRows IXLRange.RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate
    ) => this.RowsUsed(options, predicate);

    internal XLRangeRows RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeRow, Boolean> predicate = null
    )
    {
        XLRangeRows rows = new(this.Worksheet);
        Int32 rowCount = this.RowCount(options);

        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLRangeRow row = this.Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
            {
                rows.Add(row);
            }
        }
        return rows;
    }

    IXLRangeRows IXLRange.RowsUsed(Func<IXLRangeRow, Boolean> predicate) =>
        this.RowsUsed(predicate);

    internal XLRangeRows RowsUsed(Func<IXLRangeRow, Boolean> predicate = null) =>
        this.RowsUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLRangeColumns IXLRange.ColumnsUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate
    ) => this.ColumnsUsed(options, predicate);

    internal virtual XLRangeColumns ColumnsUsed(
        XLCellsUsedOptions options,
        Func<IXLRangeColumn, Boolean> predicate = null
    )
    {
        XLRangeColumns columns = new(this.Worksheet);
        Int32 columnCount = this.ColumnCount(options);

        for (Int32 co = 1; co <= columnCount; co++)
        {
            XLRangeColumn column = this.Column(co);

            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
            {
                columns.Add(column);
            }
        }
        return columns;
    }

    IXLRangeColumns IXLRange.ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate) =>
        this.ColumnsUsed(predicate);

    internal virtual XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, Boolean> predicate = null) =>
        this.ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);

    public XLRangeRow Row(Int32 row)
    {
        if (
            row <= 0
            || row > XlsxSharp.XLHelper.MaxRowNumber + this.RangeAddress.FirstAddress.RowNumber - 1
        )
        {
            throw new ArgumentOutOfRangeException(
                nameof(row),
                String.Format(
                    "Row number must be between 1 and {0}",
                    XlsxSharp.XLHelper.MaxRowNumber + this.RangeAddress.FirstAddress.RowNumber - 1
                )
            );
        }

        XLAddress firstCellAddress = new(
            this.Worksheet,
            this.RangeAddress.FirstAddress.RowNumber + row - 1,
            this.RangeAddress.FirstAddress.ColumnNumber,
            false,
            false
        );

        XLAddress lastCellAddress = new(
            this.Worksheet,
            this.RangeAddress.FirstAddress.RowNumber + row - 1,
            this.RangeAddress.LastAddress.ColumnNumber,
            false,
            false
        );
        return this.Worksheet.RangeRow(new XLRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XLRangeColumn Column(Int32 columnNumber)
    {
        if (
            columnNumber <= 0
            || columnNumber
                > XlsxSharp.XLHelper.MaxColumnNumber
                    + this.RangeAddress.FirstAddress.ColumnNumber
                    - 1
        )
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnNumber),
                String.Format(
                    "Column number must be between 1 and {0}",
                    XlsxSharp.XLHelper.MaxColumnNumber
                        + this.RangeAddress.FirstAddress.ColumnNumber
                        - 1
                )
            );
        }

        XLAddress firstCellAddress = new(
            this.Worksheet,
            this.RangeAddress.FirstAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
            false,
            false
        );
        XLAddress lastCellAddress = new(
            this.Worksheet,
            this.RangeAddress.LastAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
            false,
            false
        );
        return this.Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XLRangeColumn Column(String columnLetter) =>
        this.Column(XlsxSharp.XLHelper.GetColumnNumberFromLetter(columnLetter));

    internal IEnumerable<XLRange> Split(IXLRangeAddress anotherRange, bool includeIntersection)
    {
        if (!this.RangeAddress.Intersects(anotherRange))
        {
            yield return this;
            yield break;
        }

        int thisRow1 = this.RangeAddress.FirstAddress.RowNumber;
        int thisRow2 = this.RangeAddress.LastAddress.RowNumber;
        int thisColumn1 = this.RangeAddress.FirstAddress.ColumnNumber;
        int thisColumn2 = this.RangeAddress.LastAddress.ColumnNumber;

        int otherRow1 = Math.Min(
            Math.Max(thisRow1, anotherRange.FirstAddress.RowNumber),
            thisRow2 + 1
        );
        int otherRow2 = Math.Max(
            Math.Min(thisRow2, anotherRange.LastAddress.RowNumber),
            thisRow1 - 1
        );
        int otherColumn1 = Math.Min(
            Math.Max(thisColumn1, anotherRange.FirstAddress.ColumnNumber),
            thisColumn2 + 1
        );
        int otherColumn2 = Math.Max(
            Math.Min(thisColumn2, anotherRange.LastAddress.ColumnNumber),
            thisColumn1 - 1
        );

        XLRangeAddress[] candidates =
        [
            // to the top of the intersection
            new(
                new XLAddress(thisRow1, thisColumn1, false, false),
                new XLAddress(otherRow1 - 1, thisColumn2, false, false)
            ),
            // to the left of the intersection
            new(
                new XLAddress(otherRow1, thisColumn1, false, false),
                new XLAddress(otherRow2, otherColumn1 - 1, false, false)
            ),
            includeIntersection
                ? new XLRangeAddress(
                    new XLAddress(otherRow1, otherColumn1, false, false),
                    new XLAddress(otherRow2, otherColumn2, false, false)
                )
                : XLRangeAddress.Invalid,
            // to the right of the intersection
            new(
                new XLAddress(otherRow1, otherColumn2 + 1, false, false),
                new XLAddress(otherRow2, thisColumn2, false, false)
            ),
            // to the bottom of the intersection
            new(
                new XLAddress(otherRow2 + 1, thisColumn1, false, false),
                new XLAddress(thisRow2, thisColumn2, false, false)
            ),
        ];

        foreach (XLRangeAddress rangeAddress in candidates.Where(c => c.IsValid && c.IsNormalized))
        {
            yield return this.Worksheet.Range(rangeAddress);
        }
    }

    private void TransposeRange(int squareSide)
    {
        int rowOffset = this.RangeAddress.FirstAddress.RowNumber - 1;
        int colOffset = this.RangeAddress.FirstAddress.ColumnNumber - 1;
        for (int row = 1; row <= squareSide; ++row)
        {
            for (int col = row + 1; col <= squareSide; ++col)
            {
                Point oldAddress = new(row + rowOffset, col + colOffset);
                Point newAddress = new(col + colOffset, row + rowOffset);
                this.Worksheet.Internals.CellsCollection.SwapCellsContent(oldAddress, newAddress);
            }
        }
    }

    private void TransposeMerged(Int32 squareSide)
    {
        XLRange rngToTranspose = this.Worksheet.Range(
            this.RangeAddress.FirstAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber,
            this.RangeAddress.FirstAddress.RowNumber + squareSide - 1,
            this.RangeAddress.FirstAddress.ColumnNumber + squareSide - 1
        );

        foreach (
            XLRange merge in this.Worksheet.Internals.MergedRanges.Where<XLRange>(this.Contains)
        )
        {
            merge.RangeAddress = new XLRangeAddress(
                merge.RangeAddress.FirstAddress,
                rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address
            );
        }
    }

    private void MoveOrClearForTranspose(
        XLTransposeOptions transposeOption,
        int rowCount,
        int columnCount
    )
    {
        if (transposeOption == XLTransposeOptions.MoveCells)
        {
            if (rowCount > columnCount)
            {
                this.InsertColumnsAfter(false, rowCount - columnCount, false);
            }
            else if (columnCount > rowCount)
            {
                this.InsertRowsBelow(false, columnCount - rowCount, false);
            }
        }
        else
        {
            if (rowCount > columnCount)
            {
                int toMove = rowCount - columnCount;
                XLRange rngToClear = this.Worksheet.Range(
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + 1,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + toMove
                );
                rngToClear.Clear();
            }
            else if (columnCount > rowCount)
            {
                int toMove = columnCount - rowCount;
                XLRange rngToClear = this.Worksheet.Range(
                    this.RangeAddress.LastAddress.RowNumber + 1,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.RowNumber + toMove,
                    this.RangeAddress.LastAddress.ColumnNumber
                );
                rngToClear.Clear();
            }
        }
    }

    public override bool Equals(object obj)
    {
        XLRange other = obj as XLRange;
        if (other == null)
        {
            return false;
        }

        return this.RangeAddress.Equals(other.RangeAddress)
            && this.Worksheet.Equals(other.Worksheet);
    }

    public override int GetHashCode() =>
        this.RangeAddress.GetHashCode() ^ this.Worksheet.GetHashCode();

    public new IXLRange Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLRangeColumn FindColumn(Func<IXLRangeColumn, bool> predicate)
    {
        Int32 columnCount = this.ColumnCount();
        for (Int32 c = 1; c <= columnCount; c++)
        {
            XLRangeColumn column = this.Column(c);
            if (predicate == null || predicate(column))
            {
                return column;
            }
        }
        return null;
    }

    public IXLRangeRow FindRow(Func<IXLRangeRow, bool> predicate)
    {
        Int32 rowCount = this.RowCount();
        for (Int32 r = 1; r <= rowCount; r++)
        {
            XLRangeRow row = this.Row(r);
            if (predicate(row))
            {
                return row;
            }
        }
        return null;
    }

    public override string ToString()
    {
        if (this.IsEntireSheet())
        {
            return this.Worksheet.Name;
        }
        else if (this.IsEntireRow())
        {
            return String.Concat(
                this.Worksheet.Name.EscapeSheetName(),
                '!',
                this.RangeAddress.FirstAddress.RowNumber,
                ':',
                this.RangeAddress.LastAddress.RowNumber
            );
        }
        else if (this.IsEntireColumn())
        {
            return String.Concat(
                this.Worksheet.Name.EscapeSheetName(),
                '!',
                this.RangeAddress.FirstAddress.ColumnLetter,
                ':',
                this.RangeAddress.LastAddress.ColumnLetter
            );
        }
        else
        {
            return base.ToString();
        }
    }
}

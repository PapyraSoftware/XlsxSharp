#nullable disable

using System;
using System.Linq;

namespace XlsxSharp.Excel.Tables;

internal class XLTableRange : XLRange, IXLTableRange
{
    private readonly XLTable _table;
    private readonly XLRange _range;

    public XLTableRange(XLRange range, XLTable table)
        : base(range.RangeAddress, range.Style)
    {
        this._table = table;
        this._range = range;
    }

    IXLTableRow IXLTableRange.FirstRow(Func<IXLTableRow, Boolean> predicate)
    {
        return this.FirstRow(predicate);
    }

    public XLTableRow FirstRow(Func<IXLTableRow, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.FirstRow()));
        }

        Int32 rowCount = this._range.RowCount();

        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));
            if (predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.FirstRowUsed(Func<IXLTableRow, Boolean> predicate)
    {
        return this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    public XLTableRow FirstRowUsed(Func<IXLTableRow, Boolean> predicate = null)
    {
        return this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLTableRow IXLTableRange.FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate
    )
    {
        return this.FirstRowUsed(options, predicate);
    }

    internal XLTableRow FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.FirstRowUsed(options)));
        }

        Int32 rowCount = this._range.RowCount();

        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.LastRow(Func<IXLTableRow, Boolean> predicate)
    {
        return this.LastRow(predicate);
    }

    public XLTableRow LastRow(Func<IXLTableRow, Boolean> predicate = null)
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.LastRow()));
        }

        Int32 rowCount = this._range.RowCount();

        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));
            if (predicate(row))
            {
                return row;
            }
        }
        return null;
    }

    IXLTableRow IXLTableRange.LastRowUsed(Func<IXLTableRow, Boolean> predicate)
    {
        return this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    public XLTableRow LastRowUsed(Func<IXLTableRow, Boolean> predicate = null)
    {
        return this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLTableRow IXLTableRange.LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate
    )
    {
        return this.LastRowUsed(options, predicate);
    }

    internal XLTableRow LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate = null
    )
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.LastRowUsed(options)));
        }

        Int32 rowCount = this._range.RowCount();

        for (Int32 ro = rowCount; ro >= 1; ro--)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.Row(int row)
    {
        return this.Row(row);
    }

    public new XLTableRow Row(int row)
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

        return new XLTableRow(this, base.Row(row));
    }

    public IXLTableRows Rows(Func<IXLTableRow, Boolean> predicate = null)
    {
        XLTableRows retVal = new(this.Worksheet);
        Int32 rowCount = this._range.RowCount();

        for (int r = 1; r <= rowCount; r++)
        {
            XLTableRow row = this.Row(r);
            if (predicate == null || predicate(row))
            {
                retVal.Add(row);
            }
        }
        return retVal;
    }

    public new IXLTableRows Rows(int firstRow, int lastRow)
    {
        XLTableRows retVal = new(this.Worksheet);

        for (int rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
        {
            retVal.Add(this.Row(rowNumber));
        }

        return retVal;
    }

    public new IXLTableRows Rows(string rows)
    {
        XLTableRows retVal = new(this.Worksheet);
        string[] rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            String firstRow;
            String lastRow;
            if (Enumerable.Contains(tPair, ':') || Enumerable.Contains(tPair, '-'))
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
            foreach (IXLTableRow row in this.Rows(Int32.Parse(firstRow), Int32.Parse(lastRow)))
            {
                retVal.Add(row);
            }
        }
        return retVal;
    }

    IXLTableRows IXLTableRange.RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate
    )
    {
        return this.RowsUsed(options, predicate);
    }

    internal XLTableRows RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, Boolean> predicate = null
    )
    {
        XLTableRows rows = new(this.Worksheet);
        Int32 rowCount = this.RowCount();

        for (Int32 ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = this.Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
            {
                rows.Add(row);
            }
        }
        return rows;
    }

    IXLTableRows IXLTableRange.RowsUsed(Func<IXLTableRow, Boolean> predicate)
    {
        return this.RowsUsed(predicate);
    }

    public IXLTableRows RowsUsed(Func<IXLTableRow, Boolean> predicate = null)
    {
        return this.RowsUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLTable IXLTableRange.Table
    {
        get { return this._table; }
    }

    public XLTable Table
    {
        get { return this._table; }
    }

    public new IXLTableRows InsertRowsAbove(int numberOfRows)
    {
        return XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsAbove,
            this,
            numberOfRows,
            !this.Table.ShowTotalsRow
        );
    }

    public new IXLTableRows InsertRowsBelow(int numberOfRows)
    {
        return XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsBelow,
            this,
            numberOfRows,
            !this.Table.ShowTotalsRow
        );
    }

    public new IXLRangeColumn Column(String column)
    {
        if (XlsxSharp.XLHelper.IsValidColumn(column))
        {
            Int32 coNum = XlsxSharp.XLHelper.GetColumnNumberFromLetter(column);
            return coNum > this.ColumnCount()
                ? this.Column(this._table.GetFieldIndex(column) + 1)
                : this.Column(coNum);
        }

        return this.Column(this._table.GetFieldIndex(column) + 1);
    }
}

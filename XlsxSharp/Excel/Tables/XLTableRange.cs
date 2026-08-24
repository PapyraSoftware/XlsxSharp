#nullable disable

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

    IXLTableRow IXLTableRange.FirstRow(Func<IXLTableRow, bool> predicate) =>
        this.FirstRow(predicate);

    public XLTableRow FirstRow(Func<IXLTableRow, bool> predicate = null)
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.FirstRow()));
        }

        int rowCount = this._range.RowCount();

        for (int ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));
            if (predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.FirstRowUsed(Func<IXLTableRow, bool> predicate) =>
        this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);

    public XLTableRow FirstRowUsed(Func<IXLTableRow, bool> predicate = null) =>
        this.FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLTableRow IXLTableRange.FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate
    ) => this.FirstRowUsed(options, predicate);

    internal XLTableRow FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    )
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.FirstRowUsed(options)));
        }

        int rowCount = this._range.RowCount();

        for (int ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.LastRow(Func<IXLTableRow, bool> predicate) => this.LastRow(predicate);

    public XLTableRow LastRow(Func<IXLTableRow, bool> predicate = null)
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.LastRow()));
        }

        int rowCount = this._range.RowCount();

        for (int ro = rowCount; ro >= 1; ro--)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));
            if (predicate(row))
            {
                return row;
            }
        }
        return null;
    }

    IXLTableRow IXLTableRange.LastRowUsed(Func<IXLTableRow, bool> predicate) =>
        this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);

    public XLTableRow LastRowUsed(Func<IXLTableRow, bool> predicate = null) =>
        this.LastRowUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLTableRow IXLTableRange.LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate
    ) => this.LastRowUsed(options, predicate);

    internal XLTableRow LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    )
    {
        if (predicate == null)
        {
            return new XLTableRow(this, (this._range.LastRowUsed(options)));
        }

        int rowCount = this._range.RowCount();

        for (int ro = rowCount; ro >= 1; ro--)
        {
            XLTableRow row = new(this, (this._range.Row(ro)));

            if (!row.IsEmpty(options) && predicate(row))
            {
                return row;
            }
        }

        return null;
    }

    IXLTableRow IXLTableRange.Row(int row) => this.Row(row);

    public new XLTableRow Row(int row)
    {
        if (
            row <= 0
            || row > XlsxSharp.XLHelper.MaxRowNumber + this.RangeAddress.FirstAddress.RowNumber - 1
        )
        {
            throw new ArgumentOutOfRangeException(
                nameof(row),
                string.Format(
                    "Row number must be between 1 and {0}",
                    XlsxSharp.XLHelper.MaxRowNumber + this.RangeAddress.FirstAddress.RowNumber - 1
                )
            );
        }

        return new XLTableRow(this, base.Row(row));
    }

    public IXLTableRows Rows(Func<IXLTableRow, bool> predicate = null)
    {
        XLTableRows retVal = new(this.Worksheet);
        int rowCount = this._range.RowCount();

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
            string firstRow;
            string lastRow;
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
            foreach (IXLTableRow row in this.Rows(int.Parse(firstRow), int.Parse(lastRow)))
            {
                retVal.Add(row);
            }
        }
        return retVal;
    }

    IXLTableRows IXLTableRange.RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate
    ) => this.RowsUsed(options, predicate);

    internal XLTableRows RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    )
    {
        XLTableRows rows = new(this.Worksheet);
        int rowCount = this.RowCount();

        for (int ro = 1; ro <= rowCount; ro++)
        {
            XLTableRow row = this.Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
            {
                rows.Add(row);
            }
        }
        return rows;
    }

    IXLTableRows IXLTableRange.RowsUsed(Func<IXLTableRow, bool> predicate) =>
        this.RowsUsed(predicate);

    public IXLTableRows RowsUsed(Func<IXLTableRow, bool> predicate = null) =>
        this.RowsUsed(XLCellsUsedOptions.AllContents, predicate);

    IXLTable IXLTableRange.Table => this._table;

    public XLTable Table => this._table;

    public new IXLTableRows InsertRowsAbove(int numberOfRows) =>
        XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsAbove,
            this,
            numberOfRows,
            !this.Table.ShowTotalsRow
        );

    public new IXLTableRows InsertRowsBelow(int numberOfRows) =>
        XlsxSharp.XLHelper.InsertRowsWithoutEvents(
            base.InsertRowsBelow,
            this,
            numberOfRows,
            !this.Table.ShowTotalsRow
        );

    public new IXLRangeColumn Column(string column)
    {
        if (XlsxSharp.XLHelper.IsValidColumn(column))
        {
            int coNum = XlsxSharp.XLHelper.GetColumnNumberFromLetter(column);
            return coNum > this.ColumnCount()
                ? this.Column(this._table.GetFieldIndex(column) + 1)
                : this.Column(coNum);
        }

        return this.Column(this._table.GetFieldIndex(column) + 1);
    }
}

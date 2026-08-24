#nullable disable

namespace XlsxSharp.Excel.Tables;

public interface IXLTableRange : IXLRange
{
    public IXLTable Table { get; }

    public IXLTableRow FirstRow(Func<IXLTableRow, bool> predicate = null);

    public IXLTableRow FirstRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    );

    public IXLTableRow FirstRowUsed(Func<IXLTableRow, bool> predicate = null);

    public new IXLTableRows InsertRowsAbove(int numberOfRows);

    public new IXLTableRows InsertRowsBelow(int numberOfRows);

    public IXLTableRow LastRow(Func<IXLTableRow, bool> predicate = null);

    public IXLTableRow LastRowUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    );

    public IXLTableRow LastRowUsed(Func<IXLTableRow, bool> predicate = null);

    /// <summary>
    /// Rows the specified row.
    /// </summary>
    /// <param name="row">1-based row number relative to the first row of this range.</param>
    public new IXLTableRow Row(int row);

    public IXLTableRows Rows(Func<IXLTableRow, bool> predicate = null);

    /// <summary>
    /// Returns a subset of the rows
    /// </summary>
    /// <param name="firstRow">The first row to return. 1-based row number relative to the first row of this range.</param>
    /// <param name="lastRow">The last row to return. 1-based row number relative to the first row of this range.</param>
    public new IXLTableRows Rows(int firstRow, int lastRow);

    public new IXLTableRows Rows(string rows);

    public IXLTableRows RowsUsed(
        XLCellsUsedOptions options,
        Func<IXLTableRow, bool> predicate = null
    );

    public IXLTableRows RowsUsed(Func<IXLTableRow, bool> predicate = null);
}

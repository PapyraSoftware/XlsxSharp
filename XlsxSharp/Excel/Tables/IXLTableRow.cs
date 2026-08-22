#nullable disable

using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel.Tables;

public interface IXLTableRow : IXLRangeRow
{
    public IXLCell Field(int index);

    public IXLCell Field(string name);

    public new IXLTableRow Sort();

    public new IXLTableRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public new IXLTableRow RowAbove();

    public new IXLTableRow RowAbove(int step);

    public new IXLTableRow RowBelow();

    public new IXLTableRow RowBelow(int step);

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLTableRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public new IXLTableRows InsertRowsAbove(int numberOfRows);

    public new IXLTableRows InsertRowsBelow(int numberOfRows);
}

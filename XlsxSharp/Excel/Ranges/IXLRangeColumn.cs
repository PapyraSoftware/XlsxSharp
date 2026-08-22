#nullable disable

using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

public interface IXLRangeColumn : IXLRangeBase
{
    /// <summary>
    /// Gets the cell in the specified row.
    /// </summary>
    /// <param name="rowNumber">The cell's row.</param>
    public IXLCell Cell(int rowNumber);

    /// <summary>
    /// Returns the specified group of cells, separated by commas.
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
    /// </summary>
    /// <param name="cellsInColumn">The column cells to return.</param>
    public new IXLCells Cells(string cellsInColumn);

    /// <summary>
    /// Returns the specified group of cells.
    /// </summary>
    /// <param name="firstRow">The first row in the group of cells to return.</param>
    /// <param name="lastRow">The last row in the group of cells to return.</param>
    public IXLCells Cells(int firstRow, int lastRow);

    /// <summary>
    /// Inserts X number of columns to the right of this range.
    /// <para>All cells to the right of this range will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of columns to insert.</param>
    public IXLRangeColumns InsertColumnsAfter(int numberOfColumns);

    public IXLRangeColumns InsertColumnsAfter(int numberOfColumns, bool expandRange);

    /// <summary>
    /// Inserts X number of columns to the left of this range.
    /// <para>This range and all cells to the right of this range will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of columns to insert.</param>
    public IXLRangeColumns InsertColumnsBefore(int numberOfColumns);

    public IXLRangeColumns InsertColumnsBefore(int numberOfColumns, bool expandRange);

    /// <summary>
    /// Inserts X number of cells on top of this column.
    /// <para>This column and all cells below it will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of cells to insert.</param>
    public IXLCells InsertCellsAbove(int numberOfRows);

    public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange);

    /// <summary>
    /// Inserts X number of cells below this range.
    /// <para>All cells below this column will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of cells to insert.</param>
    public IXLCells InsertCellsBelow(int numberOfRows);

    public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange);

    /// <summary>
    /// Deletes this range and shifts the cells at the right.
    /// </summary>
    public void Delete();

    /// <summary>
    /// Deletes this range and shifts the surrounding cells accordingly.
    /// </summary>
    /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    public void Delete(XLShiftDeletedCells shiftDeleteCells);

    /// <summary>
    /// Gets this column's number in the range
    /// </summary>
    public int ColumnNumber();

    /// <summary>
    /// Gets this column's letter in the range
    /// </summary>
    public string ColumnLetter();

    public int CellCount();

    public IXLRangeColumn CopyTo(IXLCell target);

    public IXLRangeColumn CopyTo(IXLRangeBase target);

    public IXLRangeColumn Sort(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRangeColumn Column(int start, int end);

    public IXLRangeColumn Column(IXLCell start, IXLCell end);

    public IXLRangeColumns Columns(string columns);

    public IXLRangeColumn ColumnLeft();

    public IXLRangeColumn ColumnLeft(int step);

    public IXLRangeColumn ColumnRight();

    public IXLRangeColumn ColumnRight(int step);

    public IXLColumn WorksheetColumn();

    public IXLTable AsTable();

    public IXLTable AsTable(string name);

    public IXLTable CreateTable();

    public IXLTable CreateTable(string name);

    /// <summary>
    /// Clears the contents of this column.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
}

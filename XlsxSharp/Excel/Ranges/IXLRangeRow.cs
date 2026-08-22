#nullable disable

using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel;

public interface IXLRangeRow : IXLRangeBase
{
    /// <summary>
    /// Gets the cell in the specified column.
    /// </summary>
    /// <param name="columnNumber">The cell's column.</param>
    public IXLCell Cell(int columnNumber);

    /// <summary>
    /// Gets the cell in the specified column.
    /// </summary>
    /// <param name="columnLetter">The cell's column.</param>
    public IXLCell Cell(string columnLetter);

    /// <summary>
    /// Returns the specified group of cells, separated by commas.
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1:2,4:5")</para>
    /// </summary>
    /// <param name="cellsInRow">The row's cells to return.</param>
    public new IXLCells Cells(string cellsInRow);

    /// <summary>
    /// Returns the specified group of cells.
    /// </summary>
    /// <param name="firstColumn">The first column in the group of cells to return.</param>
    /// <param name="lastColumn">The last column in the group of cells to return.</param>
    public IXLCells Cells(int firstColumn, int lastColumn);

    /// <summary>
    /// Returns the specified group of cells.
    /// </summary>
    /// <param name="firstColumn">The first column in the group of cells to return.</param>
    /// <param name="lastColumn">The last column in the group of cells to return.</param>
    public IXLCells Cells(string firstColumn, string lastColumn);

    /// <summary>
    /// Inserts X number of cells to the right of this row.
    /// <para>All cells to the right of this row will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of cells to insert.</param>
    public IXLCells InsertCellsAfter(int numberOfColumns);

    public IXLCells InsertCellsAfter(int numberOfColumns, bool expandRange);

    /// <summary>
    /// Inserts X number of cells to the left of this row.
    /// <para>This row and all cells to the right of it will be shifted X number of columns.</para>
    /// </summary>
    /// <param name="numberOfColumns">Number of cells to insert.</param>
    public IXLCells InsertCellsBefore(int numberOfColumns);

    public IXLCells InsertCellsBefore(int numberOfColumns, bool expandRange);

    /// <summary>
    /// Inserts X number of rows on top of this row.
    /// <para>This row and all cells below it will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of rows to insert.</param>
    public IXLRangeRows InsertRowsAbove(int numberOfRows);

    public IXLRangeRows InsertRowsAbove(int numberOfRows, bool expandRange);

    /// <summary>
    /// Inserts X number of rows below this row.
    /// <para>All cells below this row will be shifted X number of rows.</para>
    /// </summary>
    /// <param name="numberOfRows">Number of rows to insert.</param>
    public IXLRangeRows InsertRowsBelow(int numberOfRows);

    public IXLRangeRows InsertRowsBelow(int numberOfRows, bool expandRange);

    /// <summary>
    /// Deletes this range and shifts the cells below.
    /// </summary>
    public void Delete();

    /// <summary>
    /// Deletes this range and shifts the surrounding cells accordingly.
    /// </summary>
    /// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    public void Delete(XLShiftDeletedCells shiftDeleteCells);

    /// <summary>
    /// Gets this row's number in the range
    /// </summary>
    public int RowNumber();

    public int CellCount();

    public IXLRangeRow CopyTo(IXLCell target);

    public IXLRangeRow CopyTo(IXLRangeBase target);

    public IXLRangeRow Sort();

    public IXLRangeRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRangeRow Row(int start, int end);

    public IXLRangeRow Row(IXLCell start, IXLCell end);

    public IXLRangeRows Rows(string rows);

    public IXLRangeRow RowAbove();

    public IXLRangeRow RowAbove(int step);

    public IXLRangeRow RowBelow();

    public IXLRangeRow RowBelow(int step);

    public IXLRow WorksheetRow();

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLRangeRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
}

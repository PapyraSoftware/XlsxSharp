#nullable disable

using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel.Rows;

public interface IXLRow : IXLRangeBase
{
    /// <summary>
    /// Gets or sets the height of this row.
    /// </summary>
    /// <value>
    /// The width of this row in points.
    /// </value>
    public double Height { get; set; }

    /// <summary>
    /// Clears the height for the row and defaults it to the spreadsheet row height.
    /// </summary>
    public void ClearHeight();

    /// <summary>
    /// Deletes this row and shifts the rows below this one accordingly.
    /// </summary>
    /// <remarks>Don't use in a loop due to poor performance. Use <see cref="IXLRange.Delete(XLShiftDeletedCells)"/> instead.</remarks>
    public void Delete();

    /// <summary>
    /// Gets this row's number
    /// </summary>
    public int RowNumber();

    /// <summary>
    /// Inserts X number of rows below this one.
    /// <para>All rows below will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfRows">The number of rows to insert.</param>
    public IXLRows InsertRowsBelow(int numberOfRows);

    /// <summary>
    /// Inserts X number of rows above this one.
    /// <para>This row and all below will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfRows">The number of rows to insert.</param>
    public IXLRows InsertRowsAbove(int numberOfRows);

    public IXLRow AdjustToContents();

    /// <summary>
    /// Adjusts the height of the row based on its contents, starting from the startColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    public IXLRow AdjustToContents(int startColumn);

    /// <summary>
    /// Adjusts the height of the row based on its contents, starting from the startColumn and ending at endColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    /// <param name="endColumn">The column to end calculating the row height.</param>
    public IXLRow AdjustToContents(int startColumn, int endColumn);

    public IXLRow AdjustToContents(double minHeight, double maxHeight);

    public IXLRow AdjustToContents(int startColumn, double minHeight, double maxHeight);

    /// <summary>
    /// Adjust height of the column according to the content of the cells.
    /// </summary>
    /// <param name="startColumn">Number of a first column whose content is considered.</param>
    /// <param name="endColumn">Number of a last column whose content is considered.</param>
    /// <param name="minHeightPt">Minimum height of adjusted column, in points.</param>
    /// <param name="maxHeightPt">Maximum height of adjusted column, in points.</param>
    public IXLRow AdjustToContents(
        int startColumn,
        int endColumn,
        double minHeightPt,
        double maxHeightPt
    );

    /// <summary>Hides this row.</summary>
    public IXLRow Hide();

    /// <summary>Unhides this row.</summary>
    public IXLRow Unhide();

    /// <summary>
    /// Gets a value indicating whether this row is hidden or not.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this row is hidden; otherwise, <c>false</c>.
    /// </value>
    public bool IsHidden { get; }

    /// <summary>
    /// Gets or sets the outline level of this row.
    /// </summary>
    /// <value>
    /// The outline level of this row.
    /// </value>
    public int OutlineLevel { get; set; }

    /// <summary>
    /// Adds this row to the next outline level (Increments the outline level for this row by 1).
    /// </summary>
    public IXLRow Group();

    /// <summary>
    /// Adds this row to the next outline level (Increments the outline level for this row by 1).
    /// </summary>
    /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
    public IXLRow Group(bool collapse);

    /// <summary>
    /// Sets outline level for this row.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLRow Group(int outlineLevel);

    /// <summary>
    /// Sets outline level for this row.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    /// <param name="collapse">If set to <c>true</c> the row will be shown collapsed.</param>
    public IXLRow Group(int outlineLevel, bool collapse);

    /// <summary>
    /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
    /// </summary>
    public IXLRow Ungroup();

    /// <summary>
    /// Adds this row to the previous outline level (decrements the outline level for this row by 1).
    /// </summary>
    /// <param name="fromAll">If set to <c>true</c> it will remove this row from all outline levels.</param>
    public IXLRow Ungroup(bool fromAll);

    /// <summary>
    /// Show this row as collapsed.
    /// </summary>
    public IXLRow Collapse();

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
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
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

    /// <summary>Expands this row (if it's collapsed).</summary>
    public IXLRow Expand();

    public int CellCount();

    public IXLRangeRow CopyTo(IXLCell cell);

    public IXLRangeRow CopyTo(IXLRangeBase range);

    public IXLRow CopyTo(IXLRow row);

    public IXLRow Sort();

    public IXLRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRangeRow Row(int start, int end);

    public IXLRangeRow Row(IXLCell start, IXLCell end);

    public IXLRangeRows Rows(string columns);

    /// <summary>
    /// Adds a horizontal page break after this row.
    /// </summary>
    public IXLRow AddHorizontalPageBreak();

    public IXLRow RowAbove();

    public IXLRow RowAbove(int step);

    public IXLRow RowBelow();

    public IXLRow RowBelow(int step);

    /// <summary>
    /// Clears the contents of this row.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
}

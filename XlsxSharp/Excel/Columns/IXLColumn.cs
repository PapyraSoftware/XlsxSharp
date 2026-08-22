#nullable disable

using XlsxSharp.Excel.Sort;

namespace XlsxSharp.Excel;

public interface IXLColumn : IXLRangeBase
{
    /// <summary>
    /// Gets or sets the width of this column in number of characters (NoC).
    /// </summary>
    /// <remarks>
    /// NoC are a non-linear units displayed as a column width in Excel, next to pixels. NoC combined with default font
    /// of the workbook can express width of the column in pixels and other units.
    /// </remarks>
    public double Width { get; set; }

    /// <summary>
    /// Deletes this column and shifts the columns at the right of this one accordingly.
    /// </summary>
    /// <remarks>Don't use in a loop due to poor performance. Use <see cref="IXLRange.Delete(XLShiftDeletedCells)"/> instead.</remarks>
    public void Delete();

    /// <summary>
    /// Gets this column's number
    /// </summary>
    public int ColumnNumber();

    /// <summary>
    /// Gets this column's letter
    /// </summary>
    public string ColumnLetter();

    /// <summary>
    /// Inserts X number of columns at the right of this one.
    /// <para>All columns at the right will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfColumns">The number of columns to insert.</param>
    public IXLColumns InsertColumnsAfter(int numberOfColumns);

    /// <summary>
    /// Inserts X number of columns at the left of this one.
    /// <para>This column and all at the right will be shifted accordingly.</para>
    /// </summary>
    /// <param name="numberOfColumns">The number of columns to insert.</param>
    public IXLColumns InsertColumnsBefore(int numberOfColumns);

    /// <summary>
    /// Gets the cell in the specified row.
    /// </summary>
    /// <param name="rowNumber">The cell's row.</param>
    public IXLCell Cell(int rowNumber);

    /// <summary>
    /// Returns the specified group of cells, separated by commas.
    /// <para>e.g. Cells("1"), Cells("1:5"), Cells("1,3:5")</para>
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
    /// Adjusts the width of the column based on its contents.
    /// </summary>
    public IXLColumn AdjustToContents();

    /// <summary>
    /// Adjusts the width of the column based on its contents, starting from the startRow.
    /// </summary>
    /// <param name="startRow">The row to start calculating the column width.</param>
    public IXLColumn AdjustToContents(int startRow);

    /// <summary>
    /// Adjusts the width of the column based on its contents, starting from the startRow and ending at endRow.
    /// </summary>
    /// <param name="startRow">The row to start calculating the column width.</param>
    /// <param name="endRow">The row to end calculating the column width.</param>
    public IXLColumn AdjustToContents(int startRow, int endRow);

    public IXLColumn AdjustToContents(double minWidth, double maxWidth);

    public IXLColumn AdjustToContents(int startRow, double minWidth, double maxWidth);

    /// <summary>
    /// Adjust width of the column according to the content of the cells.
    /// </summary>
    /// <param name="startRow">Number of a first row whose content is considered.</param>
    /// <param name="endRow">Number of a last row whose content is considered.</param>
    /// <param name="minWidth">Minimum width of adjusted column, in NoC.</param>
    /// <param name="maxWidth">Maximum width of adjusted column, in NoC.</param>
    public IXLColumn AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth);

    /// <summary>
    /// Hides this column.
    /// </summary>
    public IXLColumn Hide();

    /// <summary>Unhides this column.</summary>
    public IXLColumn Unhide();

    /// <summary>
    /// Gets a value indicating whether this column is hidden or not.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this column is hidden; otherwise, <c>false</c>.
    /// </value>
    public bool IsHidden { get; }

    /// <summary>
    /// Gets or sets the outline level of this column.
    /// </summary>
    /// <value>
    /// The outline level of this column.
    /// </value>
    public int OutlineLevel { get; set; }

    /// <summary>
    /// Adds this column to the next outline level (Increments the outline level for this column by 1).
    /// </summary>
    public IXLColumn Group();

    /// <summary>
    /// Adds this column to the next outline level (Increments the outline level for this column by 1).
    /// </summary>
    /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
    public IXLColumn Group(bool collapse);

    /// <summary>
    /// Sets outline level for this column.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLColumn Group(int outlineLevel);

    /// <summary>
    /// Sets outline level for this column.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    /// <param name="collapse">If set to <c>true</c> the column will be shown collapsed.</param>
    public IXLColumn Group(int outlineLevel, bool collapse);

    /// <summary>
    /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
    /// </summary>
    public IXLColumn Ungroup();

    /// <summary>
    /// Adds this column to the previous outline level (decrements the outline level for this column by 1).
    /// </summary>
    /// <param name="fromAll">If set to <c>true</c> it will remove this column from all outline levels.</param>
    public IXLColumn Ungroup(bool fromAll);

    /// <summary>
    /// Show this column as collapsed.
    /// </summary>
    public IXLColumn Collapse();

    /// <summary>Expands this column (if it's collapsed).</summary>
    public IXLColumn Expand();

    public int CellCount();

    public IXLRangeColumn CopyTo(IXLCell cell);

    public IXLRangeColumn CopyTo(IXLRangeBase range);

    public IXLColumn CopyTo(IXLColumn column);

    public IXLColumn Sort(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRangeColumn Column(int start, int end);

    public IXLRangeColumn Column(IXLCell start, IXLCell end);

    public IXLRangeColumns Columns(string columns);

    /// <summary>
    /// Adds a vertical page break after this column.
    /// </summary>
    public IXLColumn AddVerticalPageBreak();

    public IXLColumn ColumnLeft();

    public IXLColumn ColumnLeft(int step);

    public IXLColumn ColumnRight();

    public IXLColumn ColumnRight(int step);

    /// <summary>
    /// Clears the contents of this column.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents);
}

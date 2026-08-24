#nullable disable

namespace XlsxSharp.Excel.Rows;

public interface IXLRows : IEnumerable<IXLRow>
{
    /// <summary>
    /// Sets the height of all rows.
    /// </summary>
    /// <value>
    /// The height of all rows.
    /// </value>
    public double Height { set; }

    /// <summary>
    /// Deletes all rows and shifts the rows below them accordingly.
    /// </summary>
    public void Delete();

    /// <summary>
    /// Adjusts the height of all rows based on its contents.
    /// </summary>
    public IXLRows AdjustToContents();

    /// <summary>
    /// Adjusts the height of all rows based on its contents, starting from the startColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    public IXLRows AdjustToContents(int startColumn);

    /// <summary>
    /// Adjusts the height of all rows based on its contents, starting from the startColumn and ending at endColumn.
    /// </summary>
    /// <param name="startColumn">The column to start calculating the row height.</param>
    /// <param name="endColumn">The column to end calculating the row height.</param>
    public IXLRows AdjustToContents(int startColumn, int endColumn);

    public IXLRows AdjustToContents(double minHeight, double maxHeight);

    public IXLRows AdjustToContents(int startColumn, double minHeight, double maxHeight);

    public IXLRows AdjustToContents(
        int startColumn,
        int endColumn,
        double minHeight,
        double maxHeight
    );

    /// <summary>
    /// Hides all rows.
    /// </summary>
    public void Hide();

    /// <summary>Unhides all rows.</summary>
    public void Unhide();

    /// <summary>
    /// Increments the outline level of all rows by 1.
    /// </summary>
    public void Group();

    /// <summary>
    /// Increments the outline level of all rows by 1.
    /// </summary>
    /// <param name="collapse">If set to <c>true</c> the rows will be shown collapsed.</param>
    public void Group(bool collapse);

    /// <summary>
    /// Sets outline level for all rows.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public void Group(int outlineLevel);

    /// <summary>
    /// Sets outline level for all rows.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    /// <param name="collapse">If set to <c>true</c> the rows will be shown collapsed.</param>
    public void Group(int outlineLevel, bool collapse);

    /// <summary>
    /// Decrements the outline level of all rows by 1.
    /// </summary>
    public void Ungroup();

    /// <summary>
    /// Decrements the outline level of all rows by 1.
    /// </summary>
    /// <param name="fromAll">If set to <c>true</c> it will remove the rows from all outline levels.</param>
    public void Ungroup(bool fromAll);

    /// <summary>
    /// Show all rows as collapsed.
    /// </summary>
    public void Collapse();

    /// <summary>Expands all rows (if they're collapsed).</summary>
    public void Expand();

    /// <summary>
    /// Returns the collection of cells.
    /// </summary>
    public IXLCells Cells();

    /// <summary>
    /// Returns the collection of cells that have a value.
    /// </summary>
    public IXLCells CellsUsed();

    /// <summary>
    /// Returns the collection of cells that have a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLCells CellsUsed(XLCellsUsedOptions options);

    public IXLStyle Style { get; set; }

    /// <summary>
    /// Adds a horizontal page break after these rows.
    /// </summary>
    public IXLRows AddHorizontalPageBreaks();

    /// <summary>
    /// Clears the contents of these rows.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public void Select();
}

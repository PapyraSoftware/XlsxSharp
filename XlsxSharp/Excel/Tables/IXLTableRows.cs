#nullable disable

namespace XlsxSharp.Excel.Tables;

public interface IXLTableRows : IEnumerable<IXLTableRow>
{
    public IXLStyle Style { get; set; }

    /// <summary>
    /// Adds a table row to this group.
    /// </summary>
    /// <param name="tableRow">The row table to add.</param>
    public void Add(IXLTableRow tableRow);

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

    /// <summary>
    /// Clears the contents of these rows.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public void Delete();

    public void Select();
}

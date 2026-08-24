#nullable disable

namespace XlsxSharp.Excel;

public interface IXLRangeRows : IEnumerable<IXLRangeRow>
{
    /// <summary>
    /// Adds a row range to this group.
    /// </summary>
    /// <param name="rowRange">The row range to add.</param>
    public void Add(IXLRangeRow rowRange);

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
    /// Deletes all rows and shifts the rows below them accordingly.
    /// </summary>
    public void Delete();

    public IXLStyle Style { get; set; }

    /// <summary>
    /// Clears the contents of these rows.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public void Select();
}

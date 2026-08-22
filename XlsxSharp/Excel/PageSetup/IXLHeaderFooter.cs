#nullable disable

namespace XlsxSharp.Excel.PageSetup;

public enum XLHFMode
{
    OddPagesOnly,
    OddAndEvenPages,
    Odd,
}

public interface IXLHeaderFooter
{
    /// <summary>
    /// Gets the left header/footer item.
    /// </summary>
    public IXLHFItem Left { get; }

    /// <summary>
    /// Gets the middle header/footer item.
    /// </summary>
    public IXLHFItem Center { get; }

    /// <summary>
    /// Gets the right header/footer item.
    /// </summary>
    public IXLHFItem Right { get; }

    /// <summary>
    /// Gets the text of the specified header/footer occurrence.
    /// </summary>
    /// <param name="occurrence">The occurrence.</param>
    public string GetText(XLHFOccurrence occurrence);

    public IXLHeaderFooter Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);
}

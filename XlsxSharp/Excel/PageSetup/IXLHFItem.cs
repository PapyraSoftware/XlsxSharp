#nullable disable

using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Excel.PageSetup;

public enum XLHFPredefinedText
{
    PageNumber,
    NumberOfPages,
    Date,
    Time,
    FullPath,
    Path,
    File,
    SheetName,
}

/// <summary>
/// An enum specifying pages where the content of the <see cref="IXLHFItem"/> be displayed.
/// </summary>
public enum XLHFOccurrence
{
    /// <summary>
    /// Header or footer on the all pages.
    /// </summary>
    AllPages,

    /// <summary>
    /// Header or footer on the odd pages. It is necessary to enable <see cref="IXLPageSetup.DifferentOddEvenPagesOnHF"/>
    /// to display the different heading/footer on odd pages. If there is a different header/footer on the first
    /// page and the <see cref="IXLPageSetup.DifferentFirstPageOnHF"/> is enabled, the first page will not use the odd
    /// page setting.
    /// </summary>
    OddPages,

    /// <summary>
    /// Header or footer on the even pages. It is necessary to enable <see cref="IXLPageSetup.DifferentOddEvenPagesOnHF"/>
    /// to display the different heading/footer on even pages.
    /// </summary>
    EvenPages,

    /// <summary>
    /// Header or footer on the first page. It is necessary to enable <see cref="IXLPageSetup.DifferentFirstPageOnHF"/>
    /// to display the different heading/footer on the first page.
    /// </summary>
    FirstPage,
}

public interface IXLHFItem : IXLWithRichString
{
    /// <summary>
    /// Gets the text of the specified header/footer occurrence.
    /// </summary>
    /// <param name="occurrence">The occurrence.</param>
    public string GetText(XLHFOccurrence occurrence);

    /// <summary>
    /// Adds the given predefined text to this header/footer item.
    /// </summary>
    /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
    public IXLRichString AddText(XLHFPredefinedText predefinedText);

    /// <summary>
    /// Adds the given text to this header/footer item.
    /// </summary>
    /// <param name="text">The text to add to this header/footer item.</param>
    /// <param name="occurrence">The occurrence for the text.</param>
    public IXLRichString AddText(string text, XLHFOccurrence occurrence);

    /// <summary>
    /// Adds the given predefined text to this header/footer item.
    /// </summary>
    /// <param name="predefinedText">The predefined text to add to this header/footer item.</param>
    /// <param name="occurrence">The occurrence for the predefined text.</param>
    public IXLRichString AddText(XLHFPredefinedText predefinedText, XLHFOccurrence occurrence);

    /// <summary>Clears the text/formats of this header/footer item.</summary>
    /// <param name="occurrence">The occurrence to clear.</param>
    public void Clear(XLHFOccurrence occurrence = XLHFOccurrence.AllPages);

    public IXLRichString AddImage(
        string imagePath,
        XLHFOccurrence occurrence = XLHFOccurrence.AllPages
    );
}

#nullable disable

namespace XlsxSharp.Excel.PageSetup;

public enum XLPageOrientation
{
    Default,
    Portrait,
    Landscape,
}

public enum XLPaperSize
{
    LetterPaper = 1,
    LetterSmallPaper = 2,
    TabloidPaper = 3,
    LedgerPaper = 4,
    LegalPaper = 5,
    StatementPaper = 6,
    ExecutivePaper = 7,
    A3Paper = 8,
    A4Paper = 9,
    A4SmallPaper = 10,
    A5Paper = 11,
    B4Paper = 12,
    B5Paper = 13,
    FolioPaper = 14,
    QuartoPaper = 15,
    StandardPaper = 16,
    StandardPaper1 = 17,
    NotePaper = 18,
    No9Envelope = 19,
    No10Envelope = 20,
    No11Envelope = 21,
    No12Envelope = 22,
    No14Envelope = 23,
    CPaper = 24,
    DPaper = 25,
    EPaper = 26,
    DlEnvelope = 27,
    C5Envelope = 28,
    C3Envelope = 29,
    C4Envelope = 30,
    C6Envelope = 31,
    C65Envelope = 32,
    B4Envelope = 33,
    B5Envelope = 34,
    B6Envelope = 35,
    ItalyEnvelope = 36,
    MonarchEnvelope = 37,
    No634Envelope = 38,
    UsStandardFanfold = 39,
    GermanStandardFanfold = 40,
    GermanLegalFanfold = 41,
    IsoB4 = 42,
    JapaneseDoublePostcard = 43,
    StandardPaper2 = 44,
    StandardPaper3 = 45,
    StandardPaper4 = 46,
    InviteEnvelope = 47,
    LetterExtraPaper = 50,
    LegalExtraPaper = 51,
    TabloidExtraPaper = 52,
    A4ExtraPaper = 53,
    LetterTransversePaper = 54,
    A4TransversePaper = 55,
    LetterExtraTransversePaper = 56,
    SuperaSuperaA4Paper = 57,
    SuperbSuperbA3Paper = 58,
    LetterPlusPaper = 59,
    A4PlusPaper = 60,
    A5TransversePaper = 61,
    JisB5TransversePaper = 62,
    A3ExtraPaper = 63,
    A5ExtraPaper = 64,
    IsoB5ExtraPaper = 65,
    A2Paper = 66,
    A3TransversePaper = 67,
    A3ExtraTransversePaper = 68,
}

public enum XLPageOrderValues
{
    DownThenOver,
    OverThenDown,
}

public enum XLShowCommentsValues
{
    None,
    AtEnd,
    AsDisplayed,
}

public enum XLPrintErrorValues
{
    Blank,
    Dash,
    Displayed,
    NA,
}

public interface IXLPageSetup
{
    /// <summary>
    /// Gets an object to manage the print areas of the worksheet.
    /// </summary>
    public IXLPrintAreas PrintAreas { get; }

    /// <summary>
    /// Gets the first row that will repeat on the top of the printed pages.
    /// <para>Use SetRowsToRepeatAtTop() to set the rows that will be repeated on the top of the printed pages.</para>
    /// </summary>
    public int FirstRowToRepeatAtTop { get; }

    /// <summary>
    /// Gets the last row that will repeat on the top of the printed pages.
    /// <para>Use SetRowsToRepeatAtTop() to set the rows that will be repeated on the top of the printed pages.</para>
    /// </summary>
    public int LastRowToRepeatAtTop { get; }

    /// <summary>
    /// Sets the rows to repeat on the top of the printed pages.
    /// </summary>
    /// <param name="range">The range of rows to repeat on the top of the printed pages.</param>
    public void SetRowsToRepeatAtTop(string range);

    /// <summary>
    /// Sets the rows to repeat on the top of the printed pages.
    /// </summary>
    /// <param name="firstRowToRepeatAtTop">The first row to repeat at top.</param>
    /// <param name="lastRowToRepeatAtTop">The last row to repeat at top.</param>
    public void SetRowsToRepeatAtTop(int firstRowToRepeatAtTop, int lastRowToRepeatAtTop);

    /// <summary>Gets the first column to repeat on the left of the printed pages.</summary>
    /// <value>The first column to repeat on the left of the printed pages.</value>
    public int FirstColumnToRepeatAtLeft { get; }

    /// <summary>Gets the last column to repeat on the left of the printed pages.</summary>
    /// <value>The last column to repeat on the left of the printed pages.</value>
    public int LastColumnToRepeatAtLeft { get; }

    /// <summary>
    /// Sets the rows to repeat on the left of the printed pages.
    /// </summary>
    /// <param name="firstColumnToRepeatAtLeft">The first column to repeat at left.</param>
    /// <param name="lastColumnToRepeatAtLeft">The last column to repeat at left.</param>
    public void SetColumnsToRepeatAtLeft(
        int firstColumnToRepeatAtLeft,
        int lastColumnToRepeatAtLeft
    );

    /// <summary>
    /// Sets the rows to repeat on the left of the printed pages.
    /// </summary>
    /// <param name="range">The range of rows to repeat on the left of the printed pages.</param>
    public void SetColumnsToRepeatAtLeft(string range);

    /// <summary>Gets or sets the page orientation for printing.</summary>
    /// <value>The page orientation.</value>
    public XLPageOrientation PageOrientation { get; set; }

    /// <summary>
    /// Gets or sets the number of pages wide (horizontal) the worksheet will be printed on.
    /// <para>If you don't specify the PagesTall, Excel will adjust that value</para>
    /// <para>based on the contents of the worksheet and the PagesWide number.</para>
    /// <para>Setting this value will override the Scale value.</para>
    /// </summary>
    public int PagesWide { get; set; }

    /// <summary>
    /// Gets or sets the number of pages tall (vertical) the worksheet will be printed on.
    /// <para>If you don't specify the PagesWide, Excel will adjust that value</para>
    /// <para>based on the contents of the worksheet and the PagesTall number.</para>
    /// <para>Setting this value will override the Scale value.</para>
    /// </summary>
    public int PagesTall { get; set; }

    /// <summary>
    /// Gets or sets the scale at which the worksheet will be printed.
    /// <para>The worksheet will be printed on as many pages as necessary to print at the given scale.</para>
    /// <para>Setting this value will override the PagesWide and PagesTall values.</para>
    /// </summary>
    public int Scale { get; set; }

    /// <summary>
    /// Gets or sets the horizontal dpi for printing the worksheet.
    /// </summary>
    public int HorizontalDpi { get; set; }

    /// <summary>
    /// Gets or sets the vertical dpi for printing the worksheet.
    /// </summary>
    public int VerticalDpi { get; set; }

    /// <summary>
    /// Gets or sets the page number that will begin the printout.
    /// <para>For example, the first page of your printout could be numbered page 5.</para>
    /// </summary>
    /// <remarks>First page number can be negative, e.g. <c>-2</c>.</remarks>
    public int? FirstPageNumber { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the worksheet will be centered on the page horizontally.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the worksheet will be centered on the page horizontally; otherwise, <c>false</c>.
    /// </value>
    public bool CenterHorizontally { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the worksheet will be centered on the page vertically.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the worksheet will be centered on the page vertically; otherwise, <c>false</c>.
    /// </value>
    public bool CenterVertically { get; set; }

    /// <summary>
    /// Sets the scale at which the worksheet will be printed. This is equivalent to setting the Scale property.
    /// <para>The worksheet will be printed on as many pages as necessary to print at the given scale.</para>
    /// <para>Setting this value will override the PagesWide and PagesTall values.</para>
    /// </summary>
    /// <param name="percentageOfNormalSize">The scale at which the worksheet will be printed.</param>
    public void AdjustTo(int percentageOfNormalSize);

    /// <summary>
    /// Gets or sets the number of pages the worksheet will be printed on.
    /// <para>This is equivalent to setting both PagesWide and PagesTall properties.</para>
    /// <para>Setting this value will override the Scale value.</para>
    /// </summary>
    /// <param name="pagesWide">The pages wide.</param>
    /// <param name="pagesTall">The pages tall.</param>
    public void FitToPages(int pagesWide, int pagesTall);

    /// <summary>
    /// Gets or sets the size of the paper to print the worksheet.
    /// </summary>
    public XLPaperSize PaperSize { get; set; }

    /// <summary>
    /// Gets an object to work with the page margins.
    /// </summary>
    public IXLMargins Margins { get; }

    /// <summary>
    /// Gets an object to work with the page headers.
    /// </summary>
    public IXLHeaderFooter Header { get; }

    /// <summary>
    /// Gets an object to work with the page footers.
    /// </summary>
    public IXLHeaderFooter Footer { get; }

    /// <summary>
    /// Gets or sets a value indicating whether Excel will automatically adjust the font size to the scale of the worksheet.
    /// </summary>
    /// <value>
    /// 	<c>true</c> if Excel will automatically adjust the font size to the scale of the worksheet; otherwise, <c>false</c>.
    /// </value>
    public bool ScaleHFWithDocument { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the header and footer margins are aligned with the left and right margins of the worksheet.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the header and footer margins are aligned with the left and right margins of the worksheet; otherwise, <c>false</c>.
    /// </value>
    public bool AlignHFWithMargins { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the gridlines will be printed.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the gridlines will be printed; otherwise, <c>false</c>.
    /// </value>
    public bool ShowGridlines { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether to show row numbers and column letters/numbers.
    /// </summary>
    /// <value>
    /// 	<c>true</c> to show row numbers and column letters/numbers; otherwise, <c>false</c>.
    /// </value>
    public bool ShowRowAndColumnHeadings { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the worksheet will be printed in black and white.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the worksheet will be printed in black and white; otherwise, <c>false</c>.
    /// </value>
    public bool BlackAndWhite { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the worksheet will be printed in draft quality.
    /// </summary>
    /// <value>
    ///   <c>true</c> if the worksheet will be printed in draft quality; otherwise, <c>false</c>.
    /// </value>
    public bool DraftQuality { get; set; }

    /// <summary>
    /// Gets or sets the page order for printing.
    /// </summary>
    public XLPageOrderValues PageOrder { get; set; }

    /// <summary>
    /// Gets or sets how the comments will be printed.
    /// </summary>
    public XLShowCommentsValues ShowComments { get; set; }

    /// <summary>
    /// Gets a list with the row breaks (for printing).
    /// </summary>
    public List<int> RowBreaks { get; }

    /// <summary>
    /// Gets a list with the column breaks (for printing).
    /// </summary>
    public List<int> ColumnBreaks { get; }

    /// <summary>
    /// Adds a horizontal page break after the given row.
    /// </summary>
    /// <param name="row">The row to insert the break.</param>
    public void AddHorizontalPageBreak(int row);

    /// <summary>
    /// Adds a vertical page break after the given column.
    /// </summary>
    /// <param name="column">The column to insert the break.</param>
    public void AddVerticalPageBreak(int column);

    /// <summary>
    /// Gets or sets how error values will be printed.
    /// </summary>
    public XLPrintErrorValues PrintErrorValue { get; set; }

    public IXLPageSetup SetPageOrientation(XLPageOrientation value);
    public IXLPageSetup SetPagesWide(int value);
    public IXLPageSetup SetPagesTall(int value);
    public IXLPageSetup SetScale(int value);
    public IXLPageSetup SetHorizontalDpi(int value);
    public IXLPageSetup SetVerticalDpi(int value);

    /// <inheritdoc cref="FirstPageNumber"/>>
    /// <param name="value">First page number or <c>null</c> for auto/default page numbering.</param>
    public IXLPageSetup SetFirstPageNumber(int? value);

    public IXLPageSetup SetCenterHorizontally();
    public IXLPageSetup SetCenterHorizontally(bool value);
    public IXLPageSetup SetCenterVertically();
    public IXLPageSetup SetCenterVertically(bool value);
    public IXLPageSetup SetPaperSize(XLPaperSize value);
    public IXLPageSetup SetScaleHFWithDocument();
    public IXLPageSetup SetScaleHFWithDocument(bool value);
    public IXLPageSetup SetAlignHFWithMargins();
    public IXLPageSetup SetAlignHFWithMargins(bool value);
    public IXLPageSetup SetShowGridlines();
    public IXLPageSetup SetShowGridlines(bool value);
    public IXLPageSetup SetShowRowAndColumnHeadings();
    public IXLPageSetup SetShowRowAndColumnHeadings(bool value);
    public IXLPageSetup SetBlackAndWhite();
    public IXLPageSetup SetBlackAndWhite(bool value);
    public IXLPageSetup SetDraftQuality();
    public IXLPageSetup SetDraftQuality(bool value);
    public IXLPageSetup SetPageOrder(XLPageOrderValues value);
    public IXLPageSetup SetShowComments(XLShowCommentsValues value);
    public IXLPageSetup SetPrintErrorValue(XLPrintErrorValues value);

    /// <summary>
    /// Should sheet display header/footer values with <see cref="XLHFOccurrence.FirstPage"/>? Default is
    /// <c>false</c>.
    /// </summary>
    public bool DifferentFirstPageOnHF { get; set; }

    /// <inheritdoc cref="DifferentFirstPageOnHF"/>
    public IXLPageSetup SetDifferentFirstPageOnHF();

    /// <inheritdoc cref="DifferentFirstPageOnHF"/>
    public IXLPageSetup SetDifferentFirstPageOnHF(bool value);

    /// <summary>
    /// Should sheet display header/footer values with <see cref="XLHFOccurrence.OddPages"/> or
    /// <see cref="XLHFOccurrence.EvenPages"/>? Default is <c>false</c>.
    /// </summary>
    public bool DifferentOddEvenPagesOnHF { get; set; }

    /// <inheritdoc cref="DifferentOddEvenPagesOnHF"/>
    public IXLPageSetup SetDifferentOddEvenPagesOnHF();

    /// <inheritdoc cref="DifferentOddEvenPagesOnHF"/>
    public IXLPageSetup SetDifferentOddEvenPagesOnHF(bool value);
}

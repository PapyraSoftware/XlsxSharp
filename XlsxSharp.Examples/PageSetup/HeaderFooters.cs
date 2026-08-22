using XlsxSharp.Excel;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Examples.PageSetup;

public class HeaderFooters : IXLExample
{
    public void Create(string filePath)
    {
        using XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Headers and Footers");

        // Simple left header to be placed on all pages
        ws.PageSetup.Header.Left.AddText("Created with XlsxSharp");

        // Using various font decorations for the right header on the first page only
        // Here we show different methods for setting font decorations.

        // Set single font decorations immediately
        ws.PageSetup.Header.Right.AddText("The ", XLHFOccurrence.FirstPage).SetBold();
        ws.PageSetup.Header.Right.AddText("First ", XLHFOccurrence.FirstPage)
            .SetFontColor(XLColor.Red);

        // Use the IXLRichText returned by the AddText(...) method to later on modify the font
        IXLRichString richText = ws.PageSetup.Header.Right.AddText(
            "Colorful ",
            XLHFOccurrence.FirstPage
        );
        richText.FontColor = XLColor.Blue;
        richText.Underline = XLFontUnderlineValues.Double;

        // Set multiple font decorations chained
        ws.PageSetup.Header.Right.AddText("Page", XLHFOccurrence.FirstPage)
            .SetBold()
            .SetItalic()
            .SetFontName("Broadway");

        // Using predefined header/footer text:

        // Let's put the full path to the file on the right footer of every odd page:
        ws.PageSetup.Footer.Right.AddText(XLHFPredefinedText.FullPath, XLHFOccurrence.OddPages);

        // Let's put the current page number and total pages on the center of every footer:
        ws.PageSetup.Footer.Center.AddText(XLHFPredefinedText.PageNumber, XLHFOccurrence.AllPages);
        ws.PageSetup.Footer.Center.AddText(" / ", XLHFOccurrence.AllPages);
        ws.PageSetup.Footer.Center.AddText(
            XLHFPredefinedText.NumberOfPages,
            XLHFOccurrence.AllPages
        );

        // Needed to display different header with XLHFOccurrence.FirstPage (i.e. The first colorful page).
        ws.PageSetup.DifferentFirstPageOnHF = true;

        // Needed to display different footer with XLHFOccurrence.OddPages (i.e. the full path). The odd footer is
        // displayed only from page 3 onwards, because first page has different setting.
        ws.PageSetup.DifferentOddEvenPagesOnHF = true;

        // Don't align headers and footers with the margins
        ws.PageSetup.AlignHFWithMargins = false;

        // Don't scale headers and footers with the document
        ws.PageSetup.ScaleHFWithDocument = false;

        workbook.SaveAs(filePath);
    }
}

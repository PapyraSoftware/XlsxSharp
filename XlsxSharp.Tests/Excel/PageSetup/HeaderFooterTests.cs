using XlsxSharp.Excel;
using XlsxSharp.Excel.PageSetup;

namespace XlsxSharp.Tests.Excel.PageSetup;

public class HeaderFooterTests
{
    [Test]
    public void CanChangeWorksheetHeader()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");

        ws.PageSetup.Header.Center.AddText("Initial page header", XLHFOccurrence.EvenPages);

        MemoryStream ms = new();
        wb.SaveAs(ms, true);

        wb = new XLWorkbook(ms);
        ws = wb.Worksheets.First();

        ws.PageSetup.Header.Center.Clear();
        ws.PageSetup.Header.Center.AddText("Changed header", XLHFOccurrence.EvenPages);

        wb.SaveAs(ms, true);

        wb = new XLWorkbook(ms);
        ws = wb.Worksheets.First();

        string newHeader = ws.PageSetup.Header.Center.GetText(XLHFOccurrence.EvenPages);
        ClassicAssert.AreEqual("Changed header", newHeader);
    }

    [Test]
    [Arguments("")]
    [Arguments("&L&C&\"Arial\"&9 19-10-2017 \n&9&\"Arial\" &P    &N &R")] // https://github.com/XlsxSharp/XlsxSharp/issues/563
    public void CanSetHeaderFooter(string s)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            {
                XLHeaderFooter? header = ws.PageSetup.Header as XLHeaderFooter;
                header.SetInnerText(XLHFOccurrence.AllPages, s);
            }
        }
    }
}

namespace XlsxSharp.Tests.Excel.PageSetup;

public class PageLayoutTests
{
    [Test]
    public void FirstPageNumberCanBeNegative() =>
        TestHelper.CreateSaveLoadAssert(
            (_, ws) => ws.PageSetup.FirstPageNumber = -3,
            (_, ws) => ClassicAssert.AreEqual(-3, ws.PageSetup.FirstPageNumber),
            @"Other\PageSetup\Negative_first_page_number.xlsx"
        );
}

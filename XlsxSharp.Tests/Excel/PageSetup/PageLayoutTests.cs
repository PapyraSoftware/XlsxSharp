using NUnit.Framework;

namespace XlsxSharp.Tests.Excel.PageSetup;

[TestFixture]
public class PageLayoutTests
{
    [Test]
    public void FirstPageNumberCanBeNegative()
    {
        TestHelper.CreateSaveLoadAssert(
            (_, ws) => ws.PageSetup.FirstPageNumber = -3,
            (_, ws) => Assert.AreEqual(-3, ws.PageSetup.FirstPageNumber),
            @"Other\PageSetup\Negative_first_page_number.xlsx"
        );
    }
}

using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Misc;

public class HyperlinkTests
{
    [Test]
    public void TestHyperlinks()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
            IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

            IXLCell targetCell = ws2.Cell("A1");
            IXLRange targetRange = ws2.Range("A1", "B1");

            IXLCell linkCell1 = ws1.Cell("A1");
            linkCell1.Value = "Link to IXLCell";
            linkCell1.SetHyperlink(new XLHyperlink(targetCell));
            ClassicAssert.AreEqual("Sheet2!A1", linkCell1.GetHyperlink().InternalAddress);

            IXLCell linkRange1 = ws1.Cell("A2");
            linkRange1.Value = "Link to IXLRangeBase";
            linkRange1.SetHyperlink(new XLHyperlink(targetRange));
            ClassicAssert.AreEqual("Sheet2!A1:B1", linkRange1.GetHyperlink().InternalAddress);
        }
    }
}

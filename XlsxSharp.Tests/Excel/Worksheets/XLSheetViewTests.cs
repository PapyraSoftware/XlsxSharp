using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Worksheets;

public class XlSheetViewTests
{
    [Test]
    public void CopyWorksheetSheetViews()
    {
        using XLWorkbook wb1 = new();
        using XLWorkbook wb2 = new();

        IXLWorksheet ws1 = wb1.AddWorksheet("WS1");
        ws1.SheetView.TopLeftCellAddress = ws1.Cell("AZ2000").Address;

        IXLWorksheet ws2 = ws1.CopyTo(wb2, "WS2");

        ClassicAssert.AreEqual(ws2, ws2.SheetView.Worksheet);
        ClassicAssert.AreEqual("AZ2000", ws2.SheetView.TopLeftCellAddress.ToString());
    }

    [Test]
    public void InvalidTopLeftCell()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.AddWorksheet();
        IXLWorksheet ws2 = wb.AddWorksheet();

        ClassicAssert.Throws<ArgumentException>(() =>
            ws1.SheetView.TopLeftCellAddress = ws2.Cell("A1").Address
        );
    }

    [Test]
    public void SheetViews()
    {
        using MemoryStream ms = new();
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").Address;
            wb.SaveAs(ms);
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (XLWorkbook wb = new(ms))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ClassicAssert.AreEqual("AZ2000", ws.SheetView.TopLeftCellAddress.ToString());

            ws.SheetView.TopLeftCellAddress = ws.Cell("AZ2000").CellBelow().CellRight().Address;

            wb.Save();
        }

        ms.Seek(0, SeekOrigin.Begin);

        using (XLWorkbook wb = new(ms))
        {
            IXLWorksheet ws = wb.Worksheets.First();
            ClassicAssert.AreEqual("BA2001", ws.SheetView.TopLeftCellAddress.ToString());
        }
    }
}

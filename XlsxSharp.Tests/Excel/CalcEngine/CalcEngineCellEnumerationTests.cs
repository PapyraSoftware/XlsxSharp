using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class CalcEngineCellEnumerationTests
{
    [Test]
    public void CanEnumerateCellsOverEmptySheet()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet sheet1 = wb.AddWorksheet("Sheet1");
            wb.AddWorksheet("Sheet2");

            IXLCell cell = sheet1.FirstCell();
            cell.FormulaA1 = "=SUMIFS(Sheet2!B:B, Sheet2!C:C, 1)";

            ClassicAssert.AreEqual(0, cell.Value);
        }
    }
}

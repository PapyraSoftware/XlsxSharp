using System.IO;
using System.Threading;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Globalization;

public class GlobalizationTests
{
    [Test]
    [Arguments("A1*10", 1230d)]
    [Arguments("A1/10", 12.3)]
    [Arguments("A1&\" cells\"", "123 cells")]
    [Arguments("A1&\"000\"", "123000")]
    [Arguments("ISNUMBER(A1)", true)]
    [Arguments("ISBLANK(A1)", false)]
    [Arguments("DATE(2018,1,28)", 43128d)]
    public void LoadFormulaCachedValue(string formula, object expectedValue)
    {
        Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");

        using (MemoryStream ms = new())
        {
            using (XLWorkbook book1 = new())
            {
                IXLWorksheet sheet = book1.AddWorksheet("sheet1");
                sheet.Cell("A1").Value = 123;
                sheet.Cell("A2").FormulaA1 = formula;
                SaveOptions options = new() { EvaluateFormulasBeforeSaving = true };

                book1.SaveAs(ms, options);
            }
            ms.Position = 0;

            using (XLWorkbook book2 = new(ms))
            {
                IXLWorksheet ws = book2.Worksheet(1);
                XLCellValue storedValueA2 = ws.Cell("A2").CachedValue;
                ClassicAssert.IsFalse(ws.Cell("A2").NeedsRecalculation);
                ClassicAssert.AreEqual(expectedValue, storedValueA2);
            }
        }
    }
}

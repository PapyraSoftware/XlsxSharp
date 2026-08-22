using System.IO;
using System.Threading;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Globalization;

[TestFixture]
public class GlobalizationTests
{
    [Test]
    [TestCase("A1*10", 1230d)]
    [TestCase("A1/10", 12.3)]
    [TestCase("A1&\" cells\"", "123 cells")]
    [TestCase("A1&\"000\"", "123000")]
    [TestCase("ISNUMBER(A1)", true)]
    [TestCase("ISBLANK(A1)", false)]
    [TestCase("DATE(2018,1,28)", 43128d)]
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
                Assert.IsFalse(ws.Cell("A2").NeedsRecalculation);
                Assert.AreEqual(expectedValue, storedValueA2);
            }
        }
    }
}

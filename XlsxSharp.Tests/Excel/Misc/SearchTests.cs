using System.Globalization;
using System.IO;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Misc;

public class SearchTests
{
    [Test]
    public void TestSearch()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Misc\CellValues.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();

            IXLCells foundCells;

            foundCells = ws.Search("Initial Value");
            ClassicAssert.AreEqual(1, foundCells.Count());
            ClassicAssert.AreEqual("B2", foundCells.Single().Address.ToString());
            ClassicAssert.AreEqual("Initial Value", foundCells.Single().GetText());

            foundCells = ws.Search("Using");
            ClassicAssert.AreEqual(2, foundCells.Count());
            ClassicAssert.AreEqual("D2", foundCells.First().Address.ToString());
            ClassicAssert.AreEqual("Using Get...()", foundCells.First().GetText());
            ClassicAssert.AreEqual(2, foundCells.Count());
            ClassicAssert.AreEqual("E2", foundCells.Last().Address.ToString());
            ClassicAssert.AreEqual("Using GetValue<T>()", foundCells.Last().GetText());

            foundCells = ws.Search("1234");
            ClassicAssert.AreEqual(5, foundCells.Count());
            ClassicAssert.AreEqual(
                "B5,C5,D5,E5,F5",
                string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())
            );

            foundCells = ws.Search("Sep");
            ClassicAssert.AreEqual(1, foundCells.Count());
            ClassicAssert.AreEqual(
                "G3",
                string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())
            );

            foundCells = ws.Search("1234", CompareOptions.Ordinal, true);
            ClassicAssert.AreEqual(5, foundCells.Count());
            ClassicAssert.AreEqual(
                "B5,C5,D5,E5,F5",
                string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())
            );

            foundCells = ws.Search("test case");
            ClassicAssert.AreEqual(0, foundCells.Count());

            foundCells = ws.Search("test case", CompareOptions.OrdinalIgnoreCase);
            ClassicAssert.AreEqual(6, foundCells.Count());
        }
    }

    [Test]
    public void TestSearch2()
    {
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\Misc\Formulas.xlsx")
            )
        )
        using (XLWorkbook wb = new(stream))
        {
            IXLWorksheet ws = wb.Worksheets.First();

            IXLCells foundCells;

            foundCells = ws.Search("3");
            ClassicAssert.AreEqual(10, foundCells.Count());
            ClassicAssert.AreEqual("C2", foundCells.First().Address.ToString());

            foundCells = ws.Search("A2", CompareOptions.Ordinal, true);
            ClassicAssert.AreEqual(6, foundCells.Count());
            ClassicAssert.AreEqual(
                "C2,D2,B6,C6,D6,A11",
                string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())
            );

            foundCells = ws.Search("RC", CompareOptions.Ordinal, true);
            ClassicAssert.AreEqual(3, foundCells.Count());
            ClassicAssert.AreEqual(
                "E2,E3,E4",
                string.Join(",", foundCells.Select(c => c.Address.ToString()).ToArray())
            );
        }
    }
}

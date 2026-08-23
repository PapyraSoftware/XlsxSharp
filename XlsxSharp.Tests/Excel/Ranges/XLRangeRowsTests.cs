using System.Collections.Generic;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

internal class XlRangeRowsTests
{
    [Test]
    public void Style_sets_format_of_range_rows()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.Range("B3:C7");
        IXLRangeRows rangeRows = range.Rows("1,3-4");

        rangeRows.Style.Font.FontSize = 20;

        HashSet<string> expectedChangedCells = ["B3", "C3", "B5", "C5", "B6", "C6"];
        foreach (IXLCell cell in range.Grow().Cells())
        {
            string? address = cell.Address.ToString();
            int fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            ClassicAssert.AreEqual(fontSize, cell.Style.Font.FontSize, 0, address);
        }
    }
}

using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

internal class XLRangesTests
{
    [Test]
    public void Style_sets_format_of_ranges()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRanges ranges = ws.Ranges("B3:C4,C7:D7");

        ranges.Style.Font.FontSize = 20;

        HashSet<string> expectedChangedCells = ["B3", "C3", "B4", "C4", "C7", "D7"];
        foreach (IXLCell cell in ranges.Cells())
        {
            string? address = cell.Address.ToString();
            int fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize, 0, address);
        }
    }
}

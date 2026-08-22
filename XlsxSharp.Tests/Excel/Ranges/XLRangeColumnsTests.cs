using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

internal class XlRangeColumnsTests
{
    [Test]
    public void Style_sets_format_of_range_columns()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.Range("B3:E4");
        IXLRangeColumns rangeColumns = range.Columns("A,C-D");

        rangeColumns.Style.Font.FontSize = 20;

        HashSet<string> expectedChangedCells = ["B3", "B4", "D3", "D4", "E3", "E4"];
        foreach (IXLCell cell in range.Grow().Cells())
        {
            string? address = cell.Address.ToString();
            int fontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(fontSize, cell.Style.Font.FontSize);
        }
    }
}

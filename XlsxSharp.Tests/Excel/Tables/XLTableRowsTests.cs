using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Tests.Excel.Tables;

internal class XlTableRowsTests
{
    [Test]
    public void Style_sets_format_of_rows()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLTable table = ws.Cell("B3")
            .InsertTable(
                new[]
                {
                    new { Name = "Cake", Price = 7 },
                    new { Name = "Waffle", Price = 4 },
                    new { Name = "Croissant", Price = 5 },
                    new { Name = "Pie", Price = 9 },
                }
            );
        IXLRangeRows rows = table.Rows(x => x.Cell(1).GetText().StartsWith("C"));

        // Act
        rows.Style.Font.FontSize = 20;

        // Assert
        HashSet<string> expectedChangedCells =
        [
            "B4",
            "C4", // Cake row
            "B6",
            "C6", // Croissant row
        ];

        foreach (IXLCell cell in ws.Range("A2:D7").Cells())
        {
            string? address = cell.Address.ToString();
            int expectedFontSize = expectedChangedCells.Contains(address) ? 20 : 11;
            Assert.AreEqual(expectedFontSize, cell.Style.Font.FontSize);
        }
    }
}

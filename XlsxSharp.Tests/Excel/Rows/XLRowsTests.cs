using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Rows;

internal class XlRowsTests
{
    [Test]
    public void Style_sets_format_of_rows()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Rows("2,4-5").Style.Font.FontSize = 5;
        ws.Range("A1:A6").Style.Font.Bold = true; // Materialize formats in column A

        (int Row, int FontSize)[] expected =
        [
            (Row: 1, FontSize: 11),
            (Row: 2, FontSize: 5),
            (Row: 3, FontSize: 11),
            (Row: 4, FontSize: 5),
            (Row: 5, FontSize: 5),
            (Row: 6, FontSize: 11),
        ];
        foreach ((int row, int fontSize) in expected)
        {
            Assert.AreEqual(fontSize, ws.Row(row).Style.Font.FontSize);

            IXLCell cellWithFormat = ws.Cell("A" + row);
            Assert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            Assert.True(cellWithFormat.Style.Font.Bold);

            IXLCell nonMaterializedCell = ws.Cell("B" + row);
            Assert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }
}

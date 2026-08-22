using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Columns;

internal class XLColumnsTests
{
    [Test]
    public void Style_sets_format_of_columns()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Columns("B-C,E").Style.Font.FontSize = 5;
        ws.Range("A1:F1").Style.Font.Bold = true; // Materialize formats in row 1

        (string Column, int FontSize)[] expected =
        [
            (Column: "A", FontSize: 11),
            (Column: "B", FontSize: 5),
            (Column: "C", FontSize: 5),
            (Column: "D", FontSize: 11),
            (Column: "E", FontSize: 5),
            (Column: "F", FontSize: 11),
        ];
        foreach ((string column, int fontSize) in expected)
        {
            Assert.AreEqual(fontSize, ws.Column(column).Style.Font.FontSize);

            IXLCell cellWithFormat = ws.Cell(column + "1");
            Assert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            Assert.True(cellWithFormat.Style.Font.Bold);

            IXLCell nonMaterializedCell = ws.Cell(column + "2");
            Assert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }
}

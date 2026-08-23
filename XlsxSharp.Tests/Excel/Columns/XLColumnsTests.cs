using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Columns;

internal class XlColumnsTests
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
            ClassicAssert.AreEqual(fontSize, ws.Column(column).Style.Font.FontSize);

            IXLCell cellWithFormat = ws.Cell(column + "1");
            ClassicAssert.AreEqual(fontSize, cellWithFormat.Style.Font.FontSize);
            ClassicAssert.True(cellWithFormat.Style.Font.Bold);

            IXLCell nonMaterializedCell = ws.Cell(column + "2");
            ClassicAssert.AreEqual(fontSize, nonMaterializedCell.Style.Font.FontSize);
        }
    }
}

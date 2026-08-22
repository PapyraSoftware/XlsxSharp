using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Columns;

public class DeletingColumns : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Deleting Columns");

        ws.Row(1).InsertRowsBelow(2);

        IXLRange rng1 = ws.Range("B2:D2");
        IXLRange rng2 = ws.Range("F2:G2");
        IXLRange rng3 = ws.Range("A1:A3");
        IXLColumn col1 = ws.Column(1);

        rng1.Style.Fill.BackgroundColor = XLColor.Orange;
        rng2.Style.Fill.BackgroundColor = XLColor.Blue;
        rng3.Style.Fill.BackgroundColor = XLColor.Red;
        col1.Style.Fill.BackgroundColor = XLColor.Black;

        ws.Columns("A,C,E:H").Delete();
        ws.Cell("A2").Value = "OK";
        ws.Cell("B2").Value = "OK";

        workbook.SaveAs(filePath);
    }
}

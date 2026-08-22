using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Columns;

public class ColumnSettings : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Column Settings");

        IXLColumn col1 = ws.Column("B");
        col1.Style.Fill.BackgroundColor = XLColor.Red;
        col1.Width = 20;

        IXLColumn col2 = ws.Column(4);
        col2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
        col2.Width = 5;

        workbook.SaveAs(filePath);
    }
}

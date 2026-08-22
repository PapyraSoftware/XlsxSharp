using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Styles;

public class PurpleWorksheet : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Purple Worksheet");

        ws.Style.Fill.BackgroundColor = XLColor.Purple;

        workbook.SaveAs(filePath);
    }
}

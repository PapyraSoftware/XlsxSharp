using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Styles;

public class StyleNumberFormat : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Style NumberFormat");

        int co = 2;
        int ro = 1;

        ws.Cell(++ro, co).Value = 123456.789d;
        ws.Cell(ro, co).Style.NumberFormat.Format = "$ #,##0.00";

        ws.Cell(++ro, co).Value = 12.345d;
        ws.Cell(ro, co).Style.NumberFormat.Format = "0000";

        ws.Cell(++ro, co).Value = 12.345d;
        ws.Cell(ro, co).Style.NumberFormat.NumberFormatId = 3;

        ws.Column(co).AdjustToContents();

        workbook.SaveAs(filePath);
    }
}

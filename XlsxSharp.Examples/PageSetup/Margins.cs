using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.PageSetup;

public class Margins : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Margins");
        ws.PageSetup.Margins.Top = 1;
        ws.PageSetup.Margins.Bottom = 1.25;
        ws.PageSetup.Margins.Left = 0.5;
        ws.PageSetup.Margins.Right = 0.75;
        ws.PageSetup.Margins.Footer = 0.15;
        ws.PageSetup.Margins.Header = 0.30;

        ws.PageSetup.CenterHorizontally = true;
        ws.PageSetup.CenterVertically = true;

        workbook.SaveAs(filePath);
    }
}

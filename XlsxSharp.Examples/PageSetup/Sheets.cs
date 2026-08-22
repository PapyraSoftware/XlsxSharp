using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.PageSetup;

public class Sheets : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws1 = workbook.Worksheets.Add("Separate PrintAreas");
        ws1.PageSetup.PrintAreas.Add("A1:B2");
        ws1.PageSetup.PrintAreas.Add("D3:D5");

        IXLWorksheet ws2 = workbook.Worksheets.Add("Page Breaks");
        ws2.PageSetup.PrintAreas.Add("A1:D5");
        ws2.PageSetup.AddHorizontalPageBreak(2);
        ws2.PageSetup.AddVerticalPageBreak(2);

        workbook.SaveAs(filePath);
    }
}

using System;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.PageSetup;

public class TwoPages : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        foreach (int ro in Enumerable.Range(1, 100))
        {
            foreach (int co in Enumerable.Range(1, 10))
            {
                ws.Cell(ro, co).Value = ws.Cell(ro, co).Address.ToString();
            }
        }
        ws.PageSetup.PagesWide = 1;

        wb.SaveAs(filePath);
    }
}

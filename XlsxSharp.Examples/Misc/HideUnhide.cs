using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class HideUnhide : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Hide Rows Columns");

        ws.Columns(1, 3).Hide();
        ws.Rows(1, 3).Hide();

        ws.Column(2).Unhide();
        ws.Row(2).Unhide();

        wb.SaveAs(filePath);
    }
}

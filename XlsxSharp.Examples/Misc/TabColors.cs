using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class TabColors : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook wb = new();

        IXLWorksheet wsRed = wb.Worksheets.Add("Red").SetTabColor(XLColor.Red);

        IXLWorksheet wsAccent3 = wb
            .Worksheets.Add("Accent3")
            .SetTabColor(XLColor.FromTheme(XLThemeColor.Accent3));

        IXLWorksheet wsIndexed = wb.Worksheets.Add("Indexed");
        wsIndexed.TabColor = XLColor.FromIndex(24);

        IXLWorksheet wsArgb = wb.Worksheets.Add("Argb");
        wsArgb.TabColor = XLColor.FromArgb(23, 23, 23);

        wb.SaveAs(filePath);
    }
}

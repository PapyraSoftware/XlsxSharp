using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class AutoFilter : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("AutoFilter");
        ws.Cell("A1").Value = "Names";
        ws.Cell("A2").Value = "John";
        ws.Cell("A3").Value = "Hank";
        ws.Cell("A4").Value = "Dagny";

        ws.RangeUsed().SetAutoFilter();

        // Your can turn off the autofilter by:
        // 1) worksheet.AutoFilter.Clear()
        // 2) worksheet.SetAutoFilter(false)
        // 3) Pick any range in the worksheet and call the above methods on the range

        wb.SaveAs(filePath);
    }
}

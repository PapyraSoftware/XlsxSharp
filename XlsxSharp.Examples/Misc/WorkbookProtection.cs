using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class WorkbookProtection : IXLExample
{
    public void Create(String filePath)
    {
        using XLWorkbook wb = new();
        wb.Worksheets.Add("Workbook Protection");
        wb.Protect("Abc@123");
        wb.SaveAs(filePath);
    }
}

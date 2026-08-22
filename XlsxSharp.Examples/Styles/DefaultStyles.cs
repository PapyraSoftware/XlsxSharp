using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Styles;

public class DefaultStyles : IXLExample
{
    public void Create(String filePath)
    {
        // Create our workbook
        XLWorkbook workbook = new();

        // This worksheet will have the default style, row height, column width, and page setup
        IXLWorksheet ws1 = workbook.Worksheets.Add("Default Style");

        // Change the default row height for all new worksheets in this workbook
        workbook.RowHeight = 30;

        IXLWorksheet ws2 = workbook.Worksheets.Add("Tall Rows");

        // Create a worksheet and change the default row height
        IXLWorksheet ws3 = workbook.Worksheets.Add("Short Rows");
        ws3.RowHeight = 7.5;

        workbook.SaveAs(filePath);
    }
}

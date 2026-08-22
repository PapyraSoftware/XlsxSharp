using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class FreezePanes : IXLExample
{
    public void Create(string filePath)
    {
        using (XLWorkbook wb = new())
        {
            // Freeze rows and columns in one shot
            IXLWorksheet ws1 = wb.AddWorksheet("Freeze1");
            ws1.Cell(5, 5).SetActive();
            ws1.SheetView.Freeze(3, 3);

            // You can also be more specific on what you want to freeze
            // For example:
            IXLWorksheet ws2 = wb.AddWorksheet("FreezeRows");
            ws2.Cell(5, 5).SetActive();
            ws2.SheetView.FreezeRows(3);

            IXLWorksheet ws3 = wb.AddWorksheet("FreezeColumns");
            ws3.Cell(5, 5).SetActive();
            ws3.SheetView.FreezeColumns(3);

            IXLWorksheet wsSplit = wb.AddWorksheet("Split View");
            wsSplit.Cell(2, 2).SetActive();
            wsSplit.SheetView.SplitRow = 3;
            wsSplit.SheetView.SplitColumn = 3;

            wb.SaveAs(filePath);
        }
    }
}

using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Columns;

public class ColumnCells : IXLExample
{
    public void Create(String filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Column Cells");

        IXLColumn columnFromWorksheet = ws.Column(1);
        columnFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
        columnFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
        columnFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
        columnFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

        IXLRangeColumn columnFromRange = ws.Range("B1:B9").FirstColumn();

        columnFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
        columnFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
        columnFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
        columnFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

        workbook.SaveAs(filePath);
    }
}

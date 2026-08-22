using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Examples.Rows;

public class RowCells : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Row Cells");

        IXLRow rowFromWorksheet = ws.Row(1);
        rowFromWorksheet.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
        rowFromWorksheet.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
        rowFromWorksheet.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
        rowFromWorksheet.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

        IXLRangeRow rowFromRange = ws.Range("A2:I2").FirstRow();

        rowFromRange.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
        rowFromRange.Cells("2").Style.Fill.BackgroundColor = XLColor.Blue;
        rowFromRange.Cells("3,5:6").Style.Fill.BackgroundColor = XLColor.Red;
        rowFromRange.Cells(8, 9).Style.Fill.BackgroundColor = XLColor.Blue;

        ws.Columns().Width = 7;

        workbook.SaveAs(filePath);
    }
}

using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Ranges;

public class DefiningRanges : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Defining a Range");

        // With a string
        IXLRange range1 = ws.Range("A1:B1");
        range1.Cell(1, 1).Value = "ws.Range(\"A1:B1\").Merge()";
        range1.Merge();

        // With two XLAddresses
        IXLRange range2 = ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address);
        range2.Cell(1, 1).Value = "ws.Range(ws.Cell(2, 1).Address, ws.Cell(2, 2).Address).Merge()";
        range2.Merge();

        // With two strings
        IXLRange range4 = ws.Range("A3", "B3");
        range4.Cell(1, 1).Value = "ws.Range(\"A3\", \"B3\").Merge()";
        range4.Merge();

        // With 4 points
        IXLRange range5 = ws.Range(4, 1, 4, 2);
        range5.Cell(1, 1).Value = "ws.Range(4, 1, 4, 2).Merge()";
        range5.Merge();

        ws.Column("A").AdjustToContents();

        workbook.SaveAs(filePath);
    }
}

using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class BlankCells : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        ws.Cell(1, 1).Value = "X";
        ws.Cell(1, 1).Clear();
        wb.SaveAs(filePath);
    }
}

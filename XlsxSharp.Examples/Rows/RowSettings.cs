using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Examples.Rows;

public class RowSettings : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet ws = workbook.Worksheets.Add("Row Settings");

        IXLRow row1 = ws.Row(2);
        row1.Style.Fill.BackgroundColor = XLColor.Red;
        row1.Height = 30;

        IXLRow row2 = ws.Row(4);
        row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
        row2.Height = 3;

        workbook.SaveAs(filePath);
    }
}

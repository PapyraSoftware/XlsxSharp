using XlsxSharp.Excel;

namespace XlsxSharp.Examples;

public class HelloWorld
{
    public void Create(string filePath)
    {
        XLWorkbook workbook = new();
        IXLWorksheet worksheet = workbook.Worksheets.Add("Sample Sheet");
        worksheet.Cell("A1").Value = "Hello World!";
        workbook.SaveAs(filePath);
    }
}

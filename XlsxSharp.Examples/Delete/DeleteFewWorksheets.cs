using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Delete;

public class DeleteFewWorksheets : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            //Note: Prepare
            {
                using XLWorkbook workbook = new();
                workbook.Worksheets.Add("1");
                workbook.Worksheets.Add("2");
                workbook.Worksheets.Add("3");
                workbook.Worksheets.Add("4");
                workbook.SaveAs(tempFile);
            }

            //Note: Delate few worksheet
            {
                using XLWorkbook workbook = new(tempFile);
                workbook.Worksheets.Delete("1");
                workbook.Worksheets.Delete("2");
                workbook.SaveAs(filePath);
            }
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }
}

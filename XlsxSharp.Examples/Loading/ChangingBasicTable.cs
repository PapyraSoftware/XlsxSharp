using System.IO;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Loading;

public class ChangingBasicTable : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            new BasicTable().Create(tempFile);
            XLWorkbook workbook = new(tempFile);
            IXLWorksheet ws = workbook.Worksheet(1);

            // Change the background color of the headers
            IXLRange rngHeaders = ws.Range("B3:F3");
            rngHeaders.Style.Fill.BackgroundColor = XLColor.LightSalmon;

            // Change the date formats
            IXLRange rngDates = ws.Range("E4:E6");
            rngDates.Style.DateFormat.Format = "MM/dd/yyyy";

            // Change the income values to text
            IXLRange rngNumbers = ws.Range("F4:F6");
            foreach (IXLCell cell in rngNumbers.Cells())
            {
                string formattedString = cell.GetFormattedString();
                cell.SetValue(formattedString + " Dollars");
            }

            ws.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
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

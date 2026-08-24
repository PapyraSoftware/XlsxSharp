using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Ranges;

public class CopyingRanges : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            new BasicTable().Create(tempFile);
            XLWorkbook workbook = new(tempFile);
            IXLWorksheet ws = workbook.Worksheet(1);

            // Define a range with the data
            IXLCell firstTableCell = ws.FirstCellUsed();
            IXLCell lastTableCell = ws.LastCellUsed();
            IXLRange rngData = ws.Range(firstTableCell.Address, lastTableCell.Address);

            // Copy the table to another worksheet
            IXLWorksheet wsCopy = workbook.Worksheets.Add("Contacts Copy");
            wsCopy.Cell(1, 1).CopyFrom(rngData);

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

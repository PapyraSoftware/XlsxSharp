using System.Collections.Generic;
using System.IO;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class LambdaExpressions : IXLExample
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
            IXLCell firstDataCell = ws.Cell("B4");
            IXLCell lastDataCell = ws.LastCellUsed();
            IXLRange rngData = ws.Range(firstDataCell.Address, lastDataCell.Address);

            // Delete all rows where Outcast = false (the 3rd column). Deleting a row of a range
            // shifts the cells below it up, which is the default for rows in a range.
            foreach (IXLRangeRow row in rngData.Rows().Where(r => !r.Cell(3).GetBoolean()))
            {
                row.Delete();
            }

            // Put a light gray background to all text cells, taken from the range before the
            // styling starts.
            List<IXLCell> textCells = [.. rngData.Cells().Where(c => c.DataType == XLDataType.Text)];
            foreach (IXLCell c in textCells)
            {
                c.Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            // Put a thick border to the bottom of the table (we may have deleted the bottom cells with the border)
            rngData.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;

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

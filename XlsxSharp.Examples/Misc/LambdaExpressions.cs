using System.Collections.Generic;
using System.IO;
using System.Linq;
using MoreLinq;
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

            // Delete all rows where Outcast = false (the 3rd column)
            rngData
                .Rows() // From all rows
                .Where(r => !r.Cell(3).GetBoolean()) // where the 3rd cell of each row is false
                .ForEach(r => r.Delete()); // delete the row and shift the cells up (the default for rows in a range)

            //// Put a light gray background to all text cells
            //rngData.Cells() // From all cells
            //        .Where(c => c.DataType == XLCellValues.Text) // where the data type is Text
            //        .ForEach(c => c.Style.Fill.BackgroundColor = XLColor.LightGray); // Fill with a light gray

            IXLCells cells = rngData.Cells();
            IEnumerable<IXLCell> filtered = cells.Where(c => c.DataType == XLDataType.Text);
            List<IXLCell> list = [.. filtered];
            foreach (IXLCell c in list)
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

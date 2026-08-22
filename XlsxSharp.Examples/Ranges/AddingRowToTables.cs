using System.IO;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Examples.Ranges;

public class AddingRowToTables : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            new BasicTable().Create(tempFile);
            XLWorkbook wb = new(tempFile);
            IXLWorksheet ws = wb.Worksheets.First();

            IXLCell firstCell = ws.FirstCellUsed();
            IXLCell lastCell = ws.LastCellUsed();
            IXLRange range = ws.Range(firstCell.Address, lastCell.Address);
            range.FirstRow().Delete(); // Deleting the "Contacts" header (we don't need it for our purposes)

            // We want to use a theme for table, not the hard coded format of the BasicTable
            range.Clear(XLClearOptions.AllFormats);
            // Put back the date and number formats
            range.Column(4).Style.NumberFormat.NumberFormatId = 15;
            range.Column(5).Style.NumberFormat.Format = "$ #,##0";

            IXLTable table = range.CreateTable(); // You can also use range.AsTable() if you want to

            ws.Cell("Q6000").Value = "dummy value";

            IXLTableRow row = table.DataRange.InsertRowsBelow(1).First();

            wb.SaveAs(filePath);
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

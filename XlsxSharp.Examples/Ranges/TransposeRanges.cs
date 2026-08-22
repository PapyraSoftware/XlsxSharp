using System.IO;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Ranges;

public class TransposeRanges : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            new BasicTable().Create(tempFile);
            XLWorkbook workbook = new(tempFile);

            IXLWorksheet ws = workbook.Worksheet(1);

            IXLRange rngTable = ws.Range("B2:F6");

            rngTable.Transpose(XLTransposeOptions.MoveCells);

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

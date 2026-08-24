using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Ranges;

public class TransposeRangesPlus : IXLExample
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

            rngTable.Row(rngTable.RowCount() - 1).Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Place some markers
            IXLCell cellNextRow = ws.Cell(
                rngTable.RangeAddress.LastAddress.RowNumber + 1,
                rngTable.RangeAddress.LastAddress.ColumnNumber
            );
            cellNextRow.Value = "ColumnRight Row";
            IXLCell cellNextColumn = ws.Cell(
                rngTable.RangeAddress.LastAddress.RowNumber,
                rngTable.RangeAddress.LastAddress.ColumnNumber + 1
            );
            cellNextColumn.Value = "ColumnRight Column";

            rngTable.Transpose(XLTransposeOptions.MoveCells);
            rngTable.Transpose(XLTransposeOptions.MoveCells);
            rngTable.Transpose(XLTransposeOptions.ReplaceCells);
            rngTable.Transpose(XLTransposeOptions.ReplaceCells);

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

using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class MergeMoves : IXLExample
{
    public void Create(string filePath)
    {
        string tempFile = ExampleHelper.GetTempFilePath(filePath);
        try
        {
            new MergeCells().Create(tempFile);
            XLWorkbook workbook = new(tempFile);

            IXLWorksheet ws = workbook.Worksheet(1);

            // Inserts don't cross any merged cells. The insert moves them and delete moves them back
            ws.Range("B1:F1").InsertRowsBelow(1);
            ws.Range("A3:A9").InsertColumnsAfter(1);
            ws.Row(1).Delete();
            ws.Column(1).Delete();

            // First insert breaks merged in column F and then re-merges it
            ws.Range("E8:E9").InsertColumnsAfter(1);
            ws.Range("F2:F8").Merge();
            ws.Range("E3:E4").InsertColumnsAfter(1);
            ws.Range("F2:F8").Merge();

            // The insert at row 2 is a master cell of merged region, thus shift will break
            // merged cells and shifts text of master cell to G and remerges shifted text
            ws.Range("E1:E2").InsertColumnsAfter(1);
            ws.Range("G2:G8").Merge();

            // Partial shifts break merged regions
            ws.Range("E1:E2").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            ws.Range("D3:E3").InsertRowsBelow(1);
            ws.Range("A1:B1").InsertRowsBelow(1);
            ws.Range("B3:D3").Merge();
            ws.Range("A1:B1").Delete(XLShiftDeletedCells.ShiftCellsUp);

            // Clear unmerges all merged regions it intersects
            ws.Range("B8:D8").Merge();
            ws.Range("D8:D9").Clear();

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

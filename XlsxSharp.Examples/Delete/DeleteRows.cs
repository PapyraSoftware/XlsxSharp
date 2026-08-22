using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Examples.Delete;

public class DeleteRows : IXLExample
{
    public void Create(string filePath)
    {
        #region Create case
        {
            using XLWorkbook workbook = new();
            IXLWorksheet ws = workbook.Worksheets.Add("Delete red rows");

            // Put a value in a few cells
            foreach (int r in Enumerable.Range(1, 5))
            foreach (int c in Enumerable.Range(1, 5))
            {
                ws.Cell(r, c).Value = $"R{r}C{c}";
            }

            IXLRows blueRow = ws.Rows(1, 2);
            IXLRow redRow = ws.Row(5);

            blueRow.Style.Fill.BackgroundColor = XLColor.Blue;

            redRow.Style.Fill.BackgroundColor = XLColor.Red;
            workbook.SaveAs(filePath);
        }
        #endregion

        #region Remove rows
        {
            using XLWorkbook workbook = new(filePath);
            IXLWorksheet ws = workbook.Worksheets.Worksheet("Delete red rows");

            ws.Rows(1, 2).Delete();
            workbook.Save();
        }
        #endregion
    }
}

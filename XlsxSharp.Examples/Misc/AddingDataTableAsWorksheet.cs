using System.Data;
using XlsxSharp.Excel;

namespace XlsxSharp.Examples.Misc;

public class AddingDataTableAsWorksheet : IXLExample
{
    public void Create(string filePath)
    {
        XLWorkbook wb = new();

        DataTable dataTable = GetTable("Information");

        // Add a DataTable as a worksheet
        wb.Worksheets.Add(dataTable);
        wb.Worksheets.First().Columns().AdjustToContents();

        wb.SaveAs(filePath);
    }

    private static DataTable GetTable(string tableName)
    {
        DataTable table = new();
        table.TableName = tableName;
        table.Columns.Add("Dosage", typeof(int));
        table.Columns.Add("Drug", typeof(string));
        table.Columns.Add("Patient", typeof(string));
        table.Columns.Add("Date", typeof(DateTime));

        table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
        table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
        table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
        table.Rows.Add(21, "Combivent", "Janet", new DateTime(2000, 1, 4));
        table.Rows.Add(100, "Dilantin", "Melanie", new DateTime(2000, 1, 5));
        return table;
    }
}

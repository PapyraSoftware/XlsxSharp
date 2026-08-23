using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.PivotTables;

public class XlPivotTableFiltersTests
{
    [Test]
    public void AddingAndRemovingFiltersShiftsPivotTableArea()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange data = ws.Cell("A1")
            .InsertData(
                new object[]
                {
                    ("Name", "City", "Flavor", "Value"),
                    ("Cake", "Tokyo", "Vanilla", 7),
                }
            );

        IXLPivotTable pt = ws.PivotTables.Add("pt", ws.Cell("E2"), data);

        // No filter, the table is at the original cell
        ClassicAssert.AreEqual("E2", ((XLPivotTable)pt).Area.ToString());

        pt.ReportFilters.Add("City");

        // First filter also adds divider row between filter and the table.
        ClassicAssert.AreEqual("E4", ((XLPivotTable)pt).Area.ToString());

        pt.ReportFilters.Add("Flavor");

        // When second filter is added, there is no need to add second divider row.
        ClassicAssert.AreEqual("E5", ((XLPivotTable)pt).Area.ToString());

        pt.ReportFilters.Remove("City");
        ClassicAssert.AreEqual("E4", ((XLPivotTable)pt).Area.ToString());

        pt.ReportFilters.Remove("Flavor");
        ClassicAssert.AreEqual("E2", ((XLPivotTable)pt).Area.ToString());
    }
}

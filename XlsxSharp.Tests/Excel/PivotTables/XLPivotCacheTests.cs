using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Tests.Excel.PivotTables;

[TestFixture]
public class XlPivotCacheTests
{
    [Test]
    public void FieldNamesKeepNamesEvenWhenSourceChange()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.FirstCell().InsertData(new[] { "Name", "Pie" });

        IXLPivotCache pivotCache = wb.PivotCaches.Add(range);
        ws.Cell("A1").Value = "Pastry";

        Assert.AreEqual(new[] { "Name" }, pivotCache.FieldNames);
    }

    [Test]
    public void RefreshUpdatesFieldNames()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.FirstCell().InsertData(new[] { "Name", "Pie" });

        IXLPivotCache pivotCache = wb.PivotCaches.Add(range);
        ws.Cell("A1").Value = "Pastry";
        pivotCache.Refresh();

        Assert.AreEqual(new[] { "Pastry" }, pivotCache.FieldNames);
    }

    [Test]
    public void RefreshRetainsSetOptions()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.FirstCell().InsertData(new[] { "Name", "Pie" });

        IXLPivotCache pivotCache = wb.PivotCaches.Add(range);

        pivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
        pivotCache.SaveSourceData = false;
        pivotCache.RefreshDataOnOpen = true;

        pivotCache.Refresh();

        Assert.AreEqual(XLItemsToRetain.None, pivotCache.ItemsToRetainPerField);
        Assert.AreEqual(false, pivotCache.SaveSourceData);
        Assert.AreEqual(true, pivotCache.RefreshDataOnOpen);
    }

    [Test]
    public void RefreshRenamedFieldIsRemovedFromPivotTable() =>
        // Pivot table has only field for Pastry, the dough is no longer in the pivot table after refresh
        TestHelper.CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                IXLRange range = ws.FirstCell()
                    .InsertData(new object[] { ("Pastry", "Dough"), ("Waffles", "Puff") });

                IXLTable table = range.CreateTable();

                IXLPivotTable pivotTable = ws.PivotTables.Add("pvt", ws.Cell("D1"), table);
                pivotTable.RowLabels.Add("Pastry");
                pivotTable.RowLabels.Add("Dough");
                pivotTable.Values.Add("Pastry").SetSummaryFormula(XLPivotSummary.Count);

                ws.Cell("B1").Value = "Mixture";
                pivotTable.PivotCache.Refresh();
            },
            @"Other\PivotTableReferenceFiles\RenamedFieldIsRemovedFromPivotTable-output.xlsx"
        );

    [Test]
    public void PreserveFieldStatisticsEvenWithoutSourceData() =>
        // Even though pivot table cache has no records in the workbook, it does contain
        // statistics about each field (e.g. types and min/max values). These are preserved
        // through load/save.
        // The cache fields in the file don't have any shared values or records, only stats,
        // and load/save preserves all Contains* flags and Min/Max values.
        TestHelper.LoadSaveAndCompare(
            @"Other\PivotTableReferenceFiles\PivotCacheWithoutSourceData-input.xlsx",
            @"Other\PivotTableReferenceFiles\PivotCacheWithoutSourceData-output.xlsx"
        );
}

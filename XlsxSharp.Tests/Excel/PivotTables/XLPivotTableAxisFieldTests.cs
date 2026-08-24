using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.PivotTables;

/// <summary>
/// Tests methods of interface <see cref="IXLPivotField"/> implemented through <see cref="XLPivotTableAxisField"/>.
/// </summary>
internal class XlPivotTableAxisFieldTests
{
    [Test]
    public void CustomName_can_be_changed()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLPivotTable pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range);
        IXLPivotField colorField = pt.RowLabels.Add("Color");

        colorField.SetCustomName("Changed color");

        ClassicAssert.AreEqual("Changed color", pt.RowLabels.Get(0).CustomName);
    }

    [Test]
    public void CustomName_throws_exception_when_name_is_already_used()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange range = ws.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLPivotTable pt = ws.PivotTables.Add("pt", ws.Cell("E1"), range);
        IXLPivotField idField = pt.RowLabels.Add("ID", "Custom ID");
        IXLPivotField colorField = pt.RowLabels.Add("Color");

        ArgumentException ex1 = ClassicAssert.Throws<ArgumentException>(() =>
            idField.SetCustomName("Color")
        )!;
        ClassicAssert.AreEqual(
            "Custom name 'Color' is already used by another field.",
            ex1.Message
        );
        ArgumentException? ex2 = ClassicAssert.Throws<ArgumentException>(() =>
            colorField.SetCustomName("Custom ID")
        );
        ClassicAssert.AreEqual(
            "Custom name 'Custom ID' is already used by another field.",
            ex2.Message
        );
    }
}

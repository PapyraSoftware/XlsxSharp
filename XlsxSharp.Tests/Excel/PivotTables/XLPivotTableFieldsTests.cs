using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.PivotTables;

/// <summary>
/// Test methods of interface <see cref="IXLPivotFields"/> implemented through <see cref="XLPivotTableAxis"/>.
/// </summary>
internal class XlPivotTableAxisTests
{
    #region IXLPivotFields methods

    #region Add

    [Test]
    public void Add_field_not_yet_in_table_adds_field_and_shared_items()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1").InsertData(new object[] { ("ID", "Count"), (1, 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        XLPivotTable internalPt = (XLPivotTable)pt;
        ClassicAssert.IsEmpty(internalPt.PivotFields[0].Items);

        IXLPivotField idField = pt
            .RowLabels.Add("ID", "Item ID")
            .AddSubtotal(XLSubtotalFunction.Automatic);

        ClassicAssert.AreEqual("ID", idField.SourceName);
        ClassicAssert.AreEqual("Item ID", idField.CustomName);
        ClassicAssert.AreEqual("Item ID", pt.RowLabels.Single().CustomName);

        // Adds values and default aggregation func to items of the field
        IReadOnlyList<XLPivotFieldItem> fieldItems = internalPt.PivotFields[0].Items;
        ClassicAssert.AreEqual(2, fieldItems.Count);
        ClassicAssert.AreEqual(XLPivotItemType.Data, fieldItems[0].ItemType);
        ClassicAssert.AreEqual(0, fieldItems[0].ItemIndex);
        ClassicAssert.AreEqual(XLPivotItemType.Default, fieldItems[1].ItemType);
    }

    [Test]
    public void Same_field_cant_be_added_twice_to_same_axis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1").InsertData(new object[] { ("ID", "Count"), (1, 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            pt.RowLabels.Add("ID", "Item ID")
        )!;
        ClassicAssert.AreEqual("Custom name 'Item ID' is already used.", ex.Message);
    }

    [Test]
    public void Add_field_must_exist_in_cache()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1").InsertData(new object[] { ("ID", "Count"), (1, 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        ClassicAssert.DoesNotThrow(() => pt.RowLabels.Add("ID", "Item ID"));

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            pt.RowLabels.Add("nonexistent")
        )!;
        ClassicAssert.AreEqual("Field 'nonexistent' not found in pivot cache.", ex.Message);
    }

    #endregion

    #region Clear

    [Test]
    public void Clear_removes_all_fields_from_axis()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.RowLabels.Add("Color", "Custom color");

        pt.RowLabels.Clear();

        ClassicAssert.IsEmpty(pt.RowLabels);

        // Clear should also remove custom names and axis, otherwise there are problems loading
        // file with such remains in Excel.
        XLPivotTable internalPt = (XLPivotTable)pt;
        ClassicAssert.Null(internalPt.PivotFields[0].Name);
        ClassicAssert.Null(internalPt.PivotFields[0].Axis);
        ClassicAssert.Null(internalPt.PivotFields[1].Name);
        ClassicAssert.Null(internalPt.PivotFields[1].Axis);
    }

    #endregion

    #region Contains

    [Test]
    public void Contains_checks_whether_field_is_present()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        IXLPivotField idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        ClassicAssert.True(pt.RowLabels.Contains("id"));
        ClassicAssert.True(pt.RowLabels.Contains(idField));
        ClassicAssert.False(pt.RowLabels.Contains("color"));
        ClassicAssert.False(pt.RowLabels.Contains("nonexistent"));
    }

    #endregion

    #region Get(string sourceName)

    [Test]
    public void Get_field_by_source_name()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        ClassicAssert.AreEqual("ID", pt.RowLabels.Get("id").SourceName);
        KeyNotFoundException ex = ClassicAssert.Throws<KeyNotFoundException>(() =>
            pt.RowLabels.Get("color")
        )!;
        ClassicAssert.AreEqual("Field with source name 'color' not found in AxisRow.", ex.Message);
    }

    #endregion

    #region Get(int)

    [Test]
    public void Get_field_by_index()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        ClassicAssert.AreEqual("ID", pt.RowLabels.Get(0).SourceName);
        ClassicAssert.Throws<IndexOutOfRangeException>(() => pt.RowLabels.Get(-2));
        ClassicAssert.Throws<IndexOutOfRangeException>(() => pt.RowLabels.Get(1));
    }

    #endregion

    #region IndexOf

    [Test]
    public void IndexOf_finds_field_in_axis_by_source_name()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        IXLPivotField idField = pt.RowLabels.Add("ID", "Item ID");
        pt.ColumnLabels.Add("Color");

        ClassicAssert.AreEqual(0, pt.RowLabels.IndexOf("ID"));
        ClassicAssert.AreEqual(0, pt.RowLabels.IndexOf(idField));
        ClassicAssert.AreEqual(-1, pt.RowLabels.IndexOf("item id"));
        ClassicAssert.AreEqual(-1, pt.RowLabels.IndexOf("Color"));
    }

    #endregion

    #region Remove

    [Test]
    public void Remove_removes_field()
    {
        using XLWorkbook wb = new();
        IXLWorksheet data = wb.AddWorksheet();
        IXLRange range = data.Cell("A1")
            .InsertData(new object[] { ("ID", "Color", "Count"), (1, "Blue", 10) });
        IXLWorksheet ptSheet = wb.AddWorksheet();
        IXLPivotTable pt = ptSheet.PivotTables.Add("pt", ptSheet.Cell("A1"), range);
        pt.RowLabels.Add("ID");

        pt.RowLabels.Remove("id");
        pt.RowLabels.Remove("ID"); // Doesnt throw on already removed.

        ClassicAssert.IsEmpty(pt.RowLabels);
    }

    #endregion

    #endregion
}

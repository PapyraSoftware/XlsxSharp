using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.PageSetup;

public class PageBreaksTests
{
    [Test]
    public void RowBreaksShouldBeSorted()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.PageSetup.AddHorizontalPageBreak(12);
        sheet.PageSetup.AddHorizontalPageBreak(5);
        ClassicAssert.AreEqual(new[] { 5, 10, 12 }, sheet.PageSetup.RowBreaks);
    }

    [Test]
    public void ColumnBreaksShouldBeSorted()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.PageSetup.AddVerticalPageBreak(12);
        sheet.PageSetup.AddVerticalPageBreak(5);
        ClassicAssert.AreEqual(new[] { 5, 10, 12 }, sheet.PageSetup.ColumnBreaks);
    }

    [Test]
    public void RowBreaksShiftWhenInsertedRowAbove()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(5).InsertRowsAbove(1);
        ClassicAssert.AreEqual(11, sheet.PageSetup.RowBreaks[0]);
    }

    [Test]
    public void RowBreaksNotShiftWhenInsertedRowBelow()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddHorizontalPageBreak(10);
        sheet.Row(15).InsertRowsAbove(1);
        ClassicAssert.AreEqual(10, sheet.PageSetup.RowBreaks[0]);
    }

    [Test]
    public void ColumnBreaksShiftWhenInsertedColumnBefore()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(5).InsertColumnsBefore(1);
        ClassicAssert.AreEqual(11, sheet.PageSetup.ColumnBreaks[0]);
    }

    [Test]
    public void ColumnBreaksNotShiftWhenInsertedColumnAfter()
    {
        XLWorkbook wb = new();
        IXLWorksheet sheet = wb.AddWorksheet("Sheet1");

        sheet.PageSetup.AddVerticalPageBreak(10);
        sheet.Column(15).InsertColumnsBefore(1);
        ClassicAssert.AreEqual(10, sheet.PageSetup.ColumnBreaks[0]);
    }
}

using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

public class RangeShiftingTests
{
    [Test]
    public void CellsContentShiftedAfterColumnDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            SetContent(ws.Cell("D4"));

            ws.Column("C").Delete();

            AssertContent(ws.Cell("C4"), "D4");
        }
    }

    [Test]
    public void CellsContentShiftedAfterRowDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            SetContent(ws.Cell("D4"));

            ws.Row(3).Delete();

            AssertContent(ws.Cell("D3"), "D4");
        }
    }

    [Test]
    public void CellsContentShiftedAfterColumnInserted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            SetContent(ws.Cell("D4"));

            ws.Column("C").InsertColumnsBefore(1);

            AssertContent(ws.Cell("E4"), "D4");
        }
    }

    [Test]
    public void CellsContentShiftedAfterRowInserted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            SetContent(ws.Cell("D4"));

            ws.Row(3).InsertRowsAbove(1);

            AssertContent(ws.Cell("D5"), "D4");
        }
    }

    [Test]
    public void CellsContentShiftAfterRangeDeleted()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            SetContent(ws.Cell("D4"));
            SetContent(ws.Cell("F8"));

            ws.Range("B2:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);
            ws.Range("E5:F7").Delete(XLShiftDeletedCells.ShiftCellsUp);

            AssertContent(ws.Cell("B4"), "D4");
            AssertContent(ws.Cell("F5"), "F8");
        }
    }

    [Test]
    [Arguments("A5:F5")]
    [Arguments("A5:F6")]
    public void RangesBelowStayMergedAfterRangeDeleted(string deletedRangeAddress)
    {
        //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange deletedRange = ws.Range(deletedRangeAddress);
        int rangeHeight =
            deletedRange.LastRow().RowNumber() - deletedRange.FirstRow().RowNumber() + 1;
        IXLRange mergedRange = ws.Range(
            deletedRange.LastRow().RowNumber() + 1,
            deletedRange.FirstColumn().ColumnNumber(),
            deletedRange.LastRow().RowNumber() + rangeHeight,
            deletedRange.LastColumn().ColumnNumber()
        );
        mergedRange.Merge();

        deletedRange.Delete(XLShiftDeletedCells.ShiftCellsUp);

        ClassicAssert.IsTrue(mergedRange.IsMerged());
        ClassicAssert.AreEqual(deletedRangeAddress, mergedRange.RangeAddress.ToString());
    }

    [Test]
    [Arguments("A5:A8")]
    [Arguments("A5:B8")]
    public void RangesToTheRightStayMergedAfterRangeDeleted(string deletedRangeAddress)
    {
        //There is an edge case when a merged range of same size as the deleted range got unmerged (see #2358)
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLRange deletedRange = ws.Range(deletedRangeAddress);
        int rangeWidth =
            deletedRange.LastColumn().ColumnNumber()
            - deletedRange.FirstColumn().ColumnNumber()
            + 1;
        IXLRange mergedRange = ws.Range(
            deletedRange.FirstRow().RowNumber(),
            deletedRange.LastColumn().ColumnNumber() + 1,
            deletedRange.LastRow().RowNumber(),
            deletedRange.LastColumn().ColumnNumber() + rangeWidth
        );
        mergedRange.Merge();

        deletedRange.Delete(XLShiftDeletedCells.ShiftCellsLeft);

        ClassicAssert.IsTrue(mergedRange.IsMerged());
        ClassicAssert.AreEqual(deletedRangeAddress, mergedRange.RangeAddress.ToString());
    }

    private static void SetContent(IXLCell cell)
    {
        cell.FormulaA1 = $"\"Formula \" & \"{cell.Address}\"";
        cell.Style.Fill.SetBackgroundColor(XLColor.Green);
        cell.CreateComment().AddText("Some comment " + cell.Address);
    }

    private static void AssertContent(IXLCell cell, string originalAddress)
    {
        ClassicAssert.AreEqual($"\"Formula \" & \"{originalAddress}\"", cell.FormulaA1);
        ClassicAssert.AreEqual(XLColor.Green, cell.Style.Fill.BackgroundColor);
        ClassicAssert.True(cell.HasComment);
        ClassicAssert.AreEqual($"Some comment {originalAddress}", cell.GetComment().Text);
    }
}

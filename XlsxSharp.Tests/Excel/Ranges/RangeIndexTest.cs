using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Index;
using XlsxSharp.Excel.Patterns;

namespace XlsxSharp.Tests.Excel.Ranges;

public class RangeIndexTest
{
    private const int TestCount = 10000;

    [Test]
    public void FindExistingMatches()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            for (int i = 1; i <= TestCount; i++)
            {
                for (int j = 2; j <= 4; j++)
                {
                    XLAddress address = new(ws, i * 2, j, false, false);
                    ClassicAssert.True(index.Contains(in address));
                }
            }
        }
    }

    [Test]
    public void FindNonExistingMatches()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            for (int i = 1; i <= TestCount; i++)
            {
                XLAddress address = new(ws, i * 2 + 1, 3, false, false);
                ClassicAssert.False(index.Contains(in address));
            }
        }
    }

    [Test]
    public void FindExistingIntersections()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            for (int i = 1; i <= TestCount; i++)
            {
                XLRangeAddress rangeAddress = new(
                    new XLAddress(ws, i * 2, 1 + i % 4, false, false),
                    new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false)
                );

                ClassicAssert.True(index.Intersects(in rangeAddress));
            }

            for (int i = 2; i < 4; i++)
            {
                XLRangeAddress columnAddress = XLRangeAddress.EntireColumn(ws, i);
                ClassicAssert.True(index.Intersects(in columnAddress));
            }
        }
    }

    [Test]
    public void FindNonExistingIntersections()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            for (int i = 1; i <= TestCount; i++)
            {
                XLRangeAddress rangeAddress = new(
                    new XLAddress(ws, i * 2 + 1, 1 + i % 4, false, false),
                    new XLAddress(ws, i * 2 + 1, 8 - i % 3, false, false)
                );

                ClassicAssert.False(index.Intersects(in rangeAddress));
            }

            XLRangeAddress columnAddress = XLRangeAddress.EntireColumn(ws, 1);
            ClassicAssert.False(index.Intersects(in columnAddress));
            columnAddress = XLRangeAddress.EntireColumn(ws, 5);
            ClassicAssert.False(index.Intersects(in columnAddress));
        }
    }

    [Test]
    public void FindMatchAfterColumnShifting()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            ws.Column(1).InsertColumnsBefore(1000);

            XLAddress address = new(ws, 102, 1004, false, false);

            ClassicAssert.True(index.Contains(in address));
        }
    }

    [Test]
    public void FindIntersectionsAfterColumnShifting()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            ws.Column(3).InsertColumnsBefore(2);

            XLRangeAddress rangeAddress = new(ws, "F102:E103");

            ClassicAssert.True(index.Intersects(in rangeAddress));
        }
    }

    [Test]
    public void FindMatchAfterRowShifting()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            ws.Row(10).InsertRowsBelow(3);

            XLAddress address = new(ws, 103, 4, false, false);

            ClassicAssert.True(index.Contains(in address));
        }
    }

    [Test]
    public void FindIntersectionsAfterRowShifting()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            IXLRangeIndex index = this.FillIndexWithTestData(ws);

            ws.Row(10).InsertRowsBelow(3);

            XLRangeAddress rangeAddress = new(ws, "C103:E103");

            ClassicAssert.True(index.Intersects(in rangeAddress));
        }
    }

    [Test]
    public void CreateQuadTree()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            Quadrant quadTree = new();
            XLRange? range = ws.Range("BT76:CA87");

            quadTree.Add(range);

            Quadrant level0 = quadTree;
            ClassicAssert.AreEqual(1, level0.MinimumColumn);
            ClassicAssert.AreEqual(XLHelper.MaxColumnNumber, level0.MaximumColumn);
            ClassicAssert.AreEqual(1, level0.MinimumRow);
            ClassicAssert.AreEqual(XLHelper.MaxRowNumber, level0.MaximumRow);
            ClassicAssert.IsNull(level0.Ranges);
            ClassicAssert.AreEqual(128, level0.Children.Count);
            ClassicAssert.True(level0.Children.All(child => child.Level == 1));
            ClassicAssert.AreEqual(
                64,
                level0.Children.Count(child =>
                    child.MinimumColumn == 1 && child.MaximumColumn == 8192 && child.X == 0
                )
            );
            ClassicAssert.AreEqual(
                64,
                level0.Children.Count(child =>
                    child.MinimumColumn == 8193 && child.MaximumColumn == 16384 && child.X == 1
                )
            );
            ClassicAssert.AreEqual(
                2,
                level0.Children.Count(child =>
                    child.MinimumRow == 1 && child.MaximumRow == 8192 && child.Y == 0
                )
            );
            ClassicAssert.AreEqual(
                2,
                level0.Children.Count(child =>
                    child.MinimumRow == 16385 && child.MaximumRow == 24576 && child.Y == 2
                )
            );

            ClassicAssert.True(level0.Children.ElementAt(0).Children.Any());
            ClassicAssert.True(level0.Children.Skip(1).All(child => child.Children == null));

            Quadrant level8 = level0
                .Children.First() // 1
                .Children.First() // 2
                .Children.First() // 3
                .Children.First() // 4
                .Children.First() // 5
                .Children.First() // 6
                .Children.First() // 7
                .Children.Last(); // 8

            ClassicAssert.AreEqual(65, level8.MinimumColumn);
            ClassicAssert.AreEqual(65, level8.MinimumRow);
            ClassicAssert.AreEqual(128, level8.MaximumColumn);
            ClassicAssert.AreEqual(128, level8.MaximumRow);

            Quadrant level9 = level8.Children.First();
            ClassicAssert.NotNull(level9.Ranges);
            ClassicAssert.AreEqual(range, level9.Ranges.Single());
        }
    }

    [Test]
    public void XlRangesCountChangesCorrectly()
    {
        using (XLWorkbook wb = new())
        {
            XLWorksheet? ws = wb.Worksheets.Add("Sheet1") as XLWorksheet;
            XLRange? range1 = ws.Range("A1:B2");
            XLRange? range2 = ws.Range("A2:B3");
            XLRange? range3 = ws.Range("A1:B2"); // same as range1

            XLRanges ranges = new(wb);
            ranges.Add(range1);
            ClassicAssert.AreEqual(1, ranges.Count);
            ranges.Add(range2);
            ClassicAssert.AreEqual(2, ranges.Count);
            ranges.Add(range3);
            ClassicAssert.AreEqual(2, ranges.Count);

            ClassicAssert.AreEqual(ranges.Count, ranges.Count);

            // Add many entries to activate QuadTree
            for (int i = 1; i <= TestCount; i++)
            {
                ranges.Add(ws.Range(i * 2, 2, i * 2, 4));
            }

            ClassicAssert.AreEqual(2 + TestCount, ranges.Count);

            for (int i = 1; i <= TestCount; i++)
            {
                ranges.Remove(ws.Range(i * 2, 2, i * 2, 4));
            }

            ClassicAssert.AreEqual(2, ranges.Count);

            ranges.Remove(range3);
            ClassicAssert.AreEqual(1, ranges.Count);
            ranges.Remove(range2);
            ClassicAssert.AreEqual(0, ranges.Count);
            ranges.Remove(range1);
            ClassicAssert.AreEqual(0, ranges.Count);
        }
    }

    private static IXLRangeIndex CreateRangeIndex(IXLWorksheet worksheet) =>
        new XLRangeIndex<IXLRangeBase>((XLWorksheet)worksheet);

    private IXLRangeIndex FillIndexWithTestData(IXLWorksheet worksheet)
    {
        List<IXLRange> ranges = [];
        for (int i = 1; i <= TestCount; i++)
        {
            ranges.Add(worksheet.Range(i * 2, 2, i * 2, 4));
        }

        IXLRangeIndex index = CreateRangeIndex(worksheet);
        ranges.ForEach(r => index.Add(r));
        return index;
    }
}

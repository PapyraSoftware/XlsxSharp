using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Tests.Excel.Styles;

[TestFixture]
public class StyleTests
{
    [Test]
    public void EmptyCellWithQuotePrefixNotTreatedAsEmpty()
    {
        using (MemoryStream ms = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet("Sheet1");
                ws.FirstCell().SetValue("Empty cell with quote prefix:");
                XLCell? cell = ws.FirstCell().CellRight() as XLCell;

                Assert.IsTrue(cell.IsEmpty());
                cell.Value = String.Empty;
                cell.Style.IncludeQuotePrefix = true;

                Assert.IsTrue(cell.IsEmpty());
                Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                XLCell cell = (XLCell)ws.Cell("B1");
                Assert.AreEqual(1, cell.MemorySstId);

                Assert.IsTrue(cell.IsEmpty());
                Assert.IsFalse(cell.IsEmpty(XLCellsUsedOptions.All));
            }
        }
    }

    [TestCase("A1", TestName = "First cell")]
    [TestCase("A2", TestName = "Cell from initialized row")]
    [TestCase("B1", TestName = "Cell from initialized column")]
    [TestCase("D4", TestName = "Initialized cell")]
    [TestCase("F6", TestName = "Non-initialized cell")]
    public void CellTakesWorksheetStyle(string cellAddress)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            ws.Column(2);
            ws.Row(2);
            ws.Cell("D4").Value = "Non empty";
            ws.Style.Font.SetFontName("Arial");
            ws.Style.Font.SetFontSize(9);

            IXLCell cell = ws.Cell(cellAddress);
            Assert.AreEqual("Arial", cell.Style.Font.FontName);
            Assert.AreEqual(9, cell.Style.Font.FontSize);
        }
    }

    [TestCaseSource(nameof(StylizedEntities))]
    public void WorksheetStyleAffectsAllNestedEntities(Func<IXLWorksheet, IXLStyle> getEntityStyle)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();

            ws.Style.Font.FontSize = 8;

            IXLStyle style = getEntityStyle(ws);

            Assert.AreEqual(8, style.Font.FontSize);
        }
    }

    // https://github.com/XlsxSharp/XlsxSharp/issues/1813
    [Test]
    public void RowColors()
    {
        TestHelper.CreateAndCompare(
            () =>
            {
                XLWorkbook wb = new();
                {
                    IXLWorksheet ws = wb.Worksheets.Add("Row Settings 1");
                    ws.Style.Fill.BackgroundColor = XLColor.Green;

                    IXLRow row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;
                    row1.Height = 30;

                    IXLRow row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                {
                    IXLWorksheet ws = wb.Worksheets.Add("Row Settings 2");
                    ws.Style.Fill.BackgroundColor = XLColor.Red;

                    IXLRow row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;

                    IXLRow row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                {
                    IXLWorksheet ws = wb.Worksheets.Add("Row Settings 3");
                    ws.Style.Fill.BackgroundColor = XLColor.Red;

                    IXLRow row1 = ws.Row(2);
                    row1.Style.Fill.BackgroundColor = XLColor.Red;
                    row1.Height = 30;

                    IXLRow row2 = ws.Row(4);
                    row2.Style.Fill.BackgroundColor = XLColor.DarkOrange;
                    row2.Height = 3;
                }

                return wb;
            },
            @"Other\StyleReferenceFiles\RowColors\output.xlsx"
        );
    }

    [Test]
    public void StyleForCellsWithoutExplicitlySetStyleUsesCombinationOfRowAndColumnsStyles()
    {
        // If a style for a cell hasn't been explicitly set (e.g. though `cell.Style.Font
        // .SetBold(true)`), it is not yet instantiated to save memory and the actual value
        // is determined by the column style and row style. Generally speaking, the axis that
        // had its value set explicitly has a precedence, but because we can't detect that with
        // current structure, use difference from worksheet as an indication of explicitly set
        // value instead.
        // If row and column style components differ, the cells at the cross are pinged, thus test
        // sets different components for each axis.
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLStyle rowStyle = ws.Row(4)
            .Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
            .Fill.SetBackgroundColor(XLColor.Blue)
            .SetIncludeQuotePrefix()
            .Protection.SetLocked(true);

        IXLStyle colStyle = ws.Column(2)
            .Style.Border.SetBottomBorder(XLBorderStyleValues.Double)
            .Font.SetFontName("Arial")
            .NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Precision2);

        IXLStyle crossCellStyle = ws.Cell(4, 2).Style;
        Assert.AreEqual(XLAlignmentHorizontalValues.Center, crossCellStyle.Alignment.Horizontal);
        Assert.AreEqual(XLBorderStyleValues.Double, crossCellStyle.Border.BottomBorder);
        Assert.AreEqual(XLColor.Blue, crossCellStyle.Fill.BackgroundColor);
        Assert.AreEqual(true, crossCellStyle.IncludeQuotePrefix);
        Assert.AreEqual(
            (int)XLPredefinedFormat.Number.Precision2,
            crossCellStyle.NumberFormat.NumberFormatId
        );
        Assert.AreEqual(true, crossCellStyle.Protection.Locked);

        IXLStyle rowCellStyle = ws.Cell(4, 3).Style;
        Assert.AreEqual(rowStyle, rowCellStyle);

        IXLStyle colCellStyle = ws.Cell(5, 2).Style;
        Assert.AreEqual(colStyle, colCellStyle);
    }

    [Test]
    public void StyleHasEqualityComparison()
    {
        Action<IXLStyle>[] changePropertyToNonDefault =
        [
            x => x.NumberFormat.SetFormat("0.00"),
            x => x.Font.SetFontSize(15),
            x => x.SetIncludeQuotePrefix(),
            x => x.Fill.SetPatternType(XLFillPatternValues.DarkGrid),
            x => x.Border.SetBottomBorder(XLBorderStyleValues.Thick),
            x => x.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right),
            x => x.Protection.SetHidden(),
        ];

        using XLWorkbook wb = new();
        foreach (Action<IXLStyle> changeProperty in changePropertyToNonDefault)
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLStyle lhs = ws.Cell("A1").Style;
            IXLStyle rhs = ws.Cell("A2").Style;

            Assert.AreEqual(lhs, rhs);
            changeProperty(lhs);
            Assert.AreNotEqual(lhs, rhs);
        }
    }

    [Test]
    public void StyleCanBeCopied()
    {
        Action<IXLStyle>[] changePropertyToNonDefault =
        [
            x => x.NumberFormat.SetFormat("0.00"),
            x => x.Font.SetFontSize(15),
            x => x.SetIncludeQuotePrefix(),
            x => x.Fill.SetPatternType(XLFillPatternValues.DarkGrid),
            x => x.Border.SetBottomBorder(XLBorderStyleValues.Thick),
            x => x.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right),
            x => x.Protection.SetHidden(),
        ];

        using XLWorkbook wb = new();
        foreach (Action<IXLStyle> changeProperty in changePropertyToNonDefault)
        {
            IXLWorksheet ws = wb.AddWorksheet();
            IXLStyle source = ws.Cell("A1").Style;
            IXLStyle target = ws.Cell("A2").Style;

            Assert.AreEqual(source, target);
            changeProperty(source);
            Assert.AreNotEqual(source, target);

            // Copy style
            target = source;

            Assert.AreEqual(source, target);
        }
    }

    private static IEnumerable<TestCaseData> StylizedEntities
    {
        get
        {
            string t = nameof(WorksheetStyleAffectsAllNestedEntities);
            yield return new TestCaseData(new Func<IXLWorksheet, IXLStyle>(ws => ws.Style)).SetName(
                t + ": Worksheet"
            );

            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Columns().Style)
            ).SetName(t + ": Columns()");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Columns(1, 3).Style)
            ).SetName(t + ": Columns(1, 3)");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Columns("B:F").Style)
            ).SetName(t + ": Columns(\"B:F\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Columns("B", "F").Style)
            ).SetName(t + ": Columns(\"B\", \"F\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Column(5).Style)
            ).SetName(t + ": Column(5)");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Column("D").Style)
            ).SetName(t + ": Column(\"D\")");

            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Rows().Style)
            ).SetName(t + ": Rows()");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Rows(1, 3).Style)
            ).SetName(t + ": Rows(1, 3)");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Rows("1:3").Style)
            ).SetName(t + ": Rows(\"1:3\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Row(5).Style)
            ).SetName(t + ": Row(5)");

            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Cells().Style)
            ).SetName(t + ": Cells()");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Cells("B2,D4").Style)
            ).SetName(t + ": Cells(\"B2, D4\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Cell("F6").Style)
            ).SetName(t + ": Cell(\"F6\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Cell(2, 3).Style)
            ).SetName(t + ": Cell(2, 3)");

            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Ranges("F6:H9,I8:K10").Style)
            ).SetName(t + ": Ranges(\"F6:H9,I8:K10\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Range("G8:H10").Style)
            ).SetName(t + ": Range(\"G8:H10\")");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Range("G8:H10").Column(1).Style)
            ).SetName(t + ": Range(\"G8:H10\").Column(1)");
            yield return new TestCaseData(
                new Func<IXLWorksheet, IXLStyle>(ws => ws.Range("G8:H10").Row(2).Style)
            ).SetName(t + ": Range(\"G8:H10\").Row(2)");
        }
    }
}

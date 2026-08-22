using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.Clearing;

[TestFixture]
public class ClearingTests
{
    private static XLColor backgroundColor = XLColor.LightBlue;
    private static XLColor foregroundColor = XLColor.DarkBrown;

    private static IXLWorkbook SetupWorkbook()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");

        IXLCell c = ws.FirstCell().SetValue("Hello world!");

        c.GetComment().AddText("Some comment");

        c.Style.Fill.BackgroundColor = backgroundColor;
        c.Style.Font.FontColor = foregroundColor;
        c.CreateDataValidation().Custom("B1");

        ////

        c = ws.FirstCell().CellBelow().SetFormulaA1("=LEFT(A1,5)");

        c.GetComment().AddText("Another comment");

        c.Style.Fill.BackgroundColor = backgroundColor;
        c.Style.Font.FontColor = foregroundColor;

        ////

        c = ws.FirstCell().CellBelow(2).SetValue(new DateTime(2018, 1, 15));

        c.GetComment().AddText("A date");

        c.Style.Fill.BackgroundColor = backgroundColor;
        c.Style.Font.FontColor = foregroundColor;

        ws.Column(1)
            .AddConditionalFormat()
            .WhenStartsWith("Hell")
            .Fill.SetBackgroundColor(XLColor.Red)
            .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.Blue)
            .Font.SetBold();

        Assert.AreEqual(XLDataType.Text, ws.Cell("A1").Value.Type);
        Assert.AreEqual(XLDataType.Text, ws.Cell("A2").Value.Type);
        Assert.AreEqual(XLDataType.DateTime, ws.Cell("A3").Value.Type);

        Assert.AreEqual(false, ws.Cell("A1").HasFormula);
        Assert.AreEqual(true, ws.Cell("A2").HasFormula);
        Assert.AreEqual(false, ws.Cell("A1").HasFormula);

        foreach (IXLCell cell in ws.Range("A1:A3").Cells())
        {
            Assert.AreEqual(backgroundColor, cell.Style.Fill.BackgroundColor);
            Assert.AreEqual(foregroundColor, cell.Style.Font.FontColor);
            Assert.IsTrue(ws.ConditionalFormats.Any());
            Assert.IsTrue(cell.HasComment);
        }

        Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);

        return wb;
    }

    [Test]
    public void WorksheetClearAll()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear();

            foreach (IXLCell c in ws.Range("A1:A10").Cells())
            {
                Assert.IsTrue(c.IsEmpty());
                Assert.AreEqual(XLDataType.Blank, c.DataType);
                Assert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                Assert.IsFalse(ws.ConditionalFormats.Any());
                Assert.IsFalse(c.HasComment);
                Assert.AreEqual(string.Empty, c.GetDataValidation().Value);
            }
        }
    }

    [Test]
    public void WorksheetClearContents()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.Contents);

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                Assert.AreEqual(XLDataType.Blank, ws.Cell("A1").DataType);
                Assert.IsTrue(c.IsEmpty(XLCellsUsedOptions.Contents));

                Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                Assert.IsTrue(ws.ConditionalFormats.Any());
                Assert.IsTrue(c.HasComment);
            }

            Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
        }
    }

    [Test]
    public void WorksheetClearNormalFormats()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.NormalFormats);

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                Assert.IsFalse(c.IsEmpty());
                Assert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                Assert.IsTrue(ws.ConditionalFormats.Any());
                Assert.IsTrue(c.HasComment);
            }

            Assert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            Assert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            Assert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
        }
    }

    [Test]
    public void WorksheetClearConditionalFormats()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.ConditionalFormats);

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                Assert.IsFalse(c.IsEmpty());
                Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                Assert.IsFalse(ws.ConditionalFormats.Any());
                Assert.IsTrue(c.HasComment);
            }

            Assert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            Assert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            Assert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
        }
    }

    [Test]
    public void WorksheetClearComments()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.Comments);

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                Assert.IsFalse(c.IsEmpty());
                Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                Assert.IsTrue(ws.ConditionalFormats.Any());
                Assert.IsFalse(c.HasComment);
            }

            Assert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            Assert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            Assert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            Assert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
        }
    }

    [Test]
    public void WorksheetClearDataValidation()
    {
        using (IXLWorkbook wb = SetupWorkbook())
        {
            IXLWorksheet ws = wb.Worksheets.First();

            ws.Clear(XLClearOptions.DataValidation);

            foreach (IXLCell c in ws.Range("A1:A3").Cells())
            {
                Assert.IsFalse(c.IsEmpty());
                Assert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                Assert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                Assert.IsTrue(ws.ConditionalFormats.Any());
                Assert.IsTrue(c.HasComment);
            }

            Assert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            Assert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            Assert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            Assert.AreEqual(string.Empty, ws.Cell("A1").GetDataValidation().Value);
        }
    }

    [Test]
    public void DeleteClearedCellValue()
    {
        using (MemoryStream ms = new())
        {
            using (IXLWorkbook wb = SetupWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.AreEqual("Hello world!", ws.Cell("A1").GetText());
                Assert.AreEqual(new DateTime(2018, 1, 15), ws.Cell("A3").GetDateTime());

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                ws.Clear(XLClearOptions.Contents);
                Assert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                Assert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());

                wb.Save();
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                Assert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                Assert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());
            }
        }
    }

    [TestCase(XLClearOptions.All, 2)]
    [TestCase(XLClearOptions.AllContents, 4)]
    [TestCase(XLClearOptions.AllFormats, 4)]
    [TestCase(XLClearOptions.Contents, 4)]
    [TestCase(XLClearOptions.MergedRanges, 2)]
    public void CanClearMergedRanges(XLClearOptions options, int expectedCount)
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Test");

            ws.Range("A1:C3").Merge();
            ws.Range("A4:B6").Merge();
            ws.Range("D1:F3").Merge();
            ws.Range("E4:F6").Merge();

            ws.Range("C1:D6").Clear(options);

            Assert.AreEqual(expectedCount, ws.MergedRanges.Count);
        }
    }
}

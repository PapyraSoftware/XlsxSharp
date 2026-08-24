using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.Clearing;

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

        ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A1").Value.Type);
        ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A2").Value.Type);
        ClassicAssert.AreEqual(XLDataType.DateTime, ws.Cell("A3").Value.Type);

        ClassicAssert.AreEqual(false, ws.Cell("A1").HasFormula);
        ClassicAssert.AreEqual(true, ws.Cell("A2").HasFormula);
        ClassicAssert.AreEqual(false, ws.Cell("A1").HasFormula);

        foreach (IXLCell cell in ws.Range("A1:A3").Cells())
        {
            ClassicAssert.AreEqual(backgroundColor, cell.Style.Fill.BackgroundColor);
            ClassicAssert.AreEqual(foregroundColor, cell.Style.Font.FontColor);
            ClassicAssert.IsTrue(ws.ConditionalFormats.Any());
            ClassicAssert.IsTrue(cell.HasComment);
        }

        ClassicAssert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);

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
                ClassicAssert.IsTrue(c.IsEmpty());
                ClassicAssert.AreEqual(XLDataType.Blank, c.DataType);
                ClassicAssert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                ClassicAssert.IsFalse(ws.ConditionalFormats.Any());
                ClassicAssert.IsFalse(c.HasComment);
                ClassicAssert.AreEqual(string.Empty, c.GetDataValidation().Value);
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
                ClassicAssert.AreEqual(XLDataType.Blank, ws.Cell("A1").DataType);
                ClassicAssert.IsTrue(c.IsEmpty(XLCellsUsedOptions.Contents));

                ClassicAssert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                ClassicAssert.IsTrue(ws.ConditionalFormats.Any());
                ClassicAssert.IsTrue(c.HasComment);
            }

            ClassicAssert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
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
                ClassicAssert.IsFalse(c.IsEmpty());
                ClassicAssert.AreEqual(ws.Style.Fill.BackgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(ws.Style.Font.FontColor, c.Style.Font.FontColor);
                ClassicAssert.IsTrue(ws.ConditionalFormats.Any());
                ClassicAssert.IsTrue(c.HasComment);
            }

            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            ClassicAssert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            ClassicAssert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
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
                ClassicAssert.IsFalse(c.IsEmpty());
                ClassicAssert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                ClassicAssert.IsFalse(ws.ConditionalFormats.Any());
                ClassicAssert.IsTrue(c.HasComment);
            }

            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            ClassicAssert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            ClassicAssert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
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
                ClassicAssert.IsFalse(c.IsEmpty());
                ClassicAssert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                ClassicAssert.IsTrue(ws.ConditionalFormats.Any());
                ClassicAssert.IsFalse(c.HasComment);
            }

            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            ClassicAssert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            ClassicAssert.AreEqual("B1", ws.Cell("A1").GetDataValidation().Value);
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
                ClassicAssert.IsFalse(c.IsEmpty());
                ClassicAssert.AreEqual(backgroundColor, c.Style.Fill.BackgroundColor);
                ClassicAssert.AreEqual(foregroundColor, c.Style.Font.FontColor);
                ClassicAssert.IsTrue(ws.ConditionalFormats.Any());
                ClassicAssert.IsTrue(c.HasComment);
            }

            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A1").DataType);
            ClassicAssert.AreEqual(XLDataType.Text, ws.Cell("A2").DataType);
            ClassicAssert.AreEqual(XLDataType.DateTime, ws.Cell("A3").DataType);

            ClassicAssert.AreEqual(string.Empty, ws.Cell("A1").GetDataValidation().Value);
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
                ClassicAssert.AreEqual("Hello world!", ws.Cell("A1").GetText());
                ClassicAssert.AreEqual(new DateTime(2018, 1, 15), ws.Cell("A3").GetDateTime());

                wb.SaveAs(ms);
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                ws.Clear(XLClearOptions.Contents);
                ClassicAssert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                ClassicAssert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());

                wb.Save();
            }

            using (XLWorkbook wb = new(ms))
            {
                IXLWorksheet ws = wb.Worksheets.First();
                ClassicAssert.AreEqual(Blank.Value, ws.Cell("A1").Value);
                ClassicAssert.Throws<InvalidCastException>(() => ws.Cell("A3").GetDateTime());
            }
        }
    }

    [Test]
    [Arguments(XLClearOptions.All, 2)]
    [Arguments(XLClearOptions.AllContents, 4)]
    [Arguments(XLClearOptions.AllFormats, 4)]
    [Arguments(XLClearOptions.Contents, 4)]
    [Arguments(XLClearOptions.MergedRanges, 2)]
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

            ClassicAssert.AreEqual(expectedCount, ws.MergedRanges.Count);
        }
    }
}

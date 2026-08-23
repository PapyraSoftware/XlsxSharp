using System;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Tests.Excel.Misc;

public class CopyContentsTests
{
    private static void CopyRowAsRange(
        IXLWorksheet originalSheet,
        int originalRowNumber,
        IXLWorksheet destSheet,
        int destRowNumber
    )
    {
        {
            IXLRow destinationRow = destSheet.Row(destRowNumber);
            destinationRow.Clear();

            IXLRow originalRow = originalSheet.Row(originalRowNumber);
            int columnNumber = originalRow
                .LastCellUsed(XLCellsUsedOptions.All)
                .Address.ColumnNumber;

            IXLRange originalRange = originalSheet.Range(
                originalRowNumber,
                1,
                originalRowNumber,
                columnNumber
            );
            IXLRange destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
            originalRange.CopyTo(destRange);
        }
    }

    [Test]
    public void CopyConditionalFormatsCount()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell()
            .AddConditionalFormat()
            .WhenContains("1")
            .Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.FirstCell().AsRange());
        ClassicAssert.AreEqual(2, ws.ConditionalFormats.Count());
    }

    [Test]
    public void CopyConditionalFormatsFixedNum()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)
            )
        );
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)
            )
        );
    }

    [Test]
    public void CopyConditionalFormatsFixedString()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "A";
        ws.Cell("B1").Value = "B";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)
            )
        );
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)
            )
        );
    }

    [Test]
    public void CopyConditionalFormatsFixedStringNum()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)
            )
        );
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)
            )
        );
    }

    [Test]
    public void CopyConditionalFormatsRelative()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1")
            .AddConditionalFormat()
            .WhenEquals("=B1")
            .Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula)
            )
        );
        ClassicAssert.IsTrue(
            ws.ConditionalFormats.Any(cf =>
                cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula)
            )
        );
    }

    [Test]
    public void TestRowCopyContents()
    {
        XLWorkbook workbook = new();
        IXLWorksheet originalSheet = workbook.Worksheets.Add("original");
        IXLWorksheet copyRowSheet = workbook.Worksheets.Add("copy row");
        IXLWorksheet copyRowAsRangeSheet = workbook.Worksheets.Add("copy row as range");
        IXLWorksheet copyRangeSheet = workbook.Worksheets.Add("copy range");

        originalSheet.Cell("A2").SetValue("test value");
        originalSheet.Range("A2:E2").Merge();

        {
            IXLRange originalRange = originalSheet.Range("A2:E2");
            IXLRange destinationRange = copyRangeSheet.Range("A2:E2");

            originalRange.CopyTo(destinationRange);
        }
        CopyRowAsRange(originalSheet, 2, copyRowAsRangeSheet, 3);
        {
            IXLRow originalRow = originalSheet.Row(2);
            IXLRow destinationRow = copyRowSheet.Row(2);
            copyRowSheet.Cell("G2").Value = "must be removed after copy";
            originalRow.CopyTo(destinationRow);
        }
        TestHelper.SaveWorkbook(workbook, "Misc", "CopyRowContents.xlsx");
    }

    [Test]
    public void UpdateCellsWorksheetTest()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
            ws1.Cell(1, 1).Value = "hello, world.";

            IXLWorksheet ws2 = ws1.CopyTo("Sheet2");

            ClassicAssert.AreEqual("Sheet1", ws1.FirstCell().Address.Worksheet.Name);
            ClassicAssert.AreEqual("Sheet2", ws2.FirstCell().Address.Worksheet.Name);
        }
    }

    [Test]
    public void CopyHyperlinksAmongSheets()
    {
        using XLWorkbook wb = new();
        IXLWorksheet source = wb.AddWorksheet();
        IXLWorksheet target = wb.AddWorksheet();
        source
            .Cell("A1")
            .SetValue("link")
            .CreateHyperlink()
            .SetValues("https://example.com", "Test tooltip");

        source.Cell("A1").AsRange().CopyTo(target.Cell("B7"));

        IXLCell cell = target.Cell("B7");
        ClassicAssert.True(cell.HasHyperlink);
        ClassicAssert.True(cell.GetHyperlink().IsExternal);
        ClassicAssert.AreEqual(new Uri("https://example.com"), cell.GetHyperlink().ExternalAddress);
        ClassicAssert.AreEqual("Test tooltip", cell.GetHyperlink().Tooltip);
    }
}

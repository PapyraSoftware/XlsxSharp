using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Tests.Excel.Styles;

public class NumberFormatTests
{
    [Test]
    public void PreserveCellFormat()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            DataTable table = new();
            table.Columns.Add("Date", typeof(DateTime));

            for (int i = 0; i < 10; i++)
            {
                table.Rows.Add(new DateTime(2017, 1, 1).AddMonths(i));
            }

            ws.Column(1).Style.NumberFormat.Format = "yy-MM-dd";
            ws.Cell("A1").InsertData(table);
            Assert.AreEqual("yy-MM-dd", ws.Cell("A5").Style.DateFormat.Format);

            ws.Row(1).Style.NumberFormat.Format = "yy-MM-dd";
            ws.Cell("A1").InsertData(table.Rows, true);
            Assert.AreEqual("yy-MM-dd", ws.Cell("E1").Style.DateFormat.Format);
        }
    }

    [Test]
    public void TestExcelNumberFormats()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");
            IXLCell c = ws.FirstCell().SetValue((41573.875));

            c.Style.NumberFormat.SetFormat("m/d/yy\\ h:mm;@");

            Assert.AreEqual("10/26/13 21:00", c.GetFormattedString());
        }
    }

    [Test]
    [SetCulture("en-US")]
    public void CellValueIsFormattedByCurrentCultureUnlessSpecifiedOtherwise()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLCell cell = ws.Cell("A1").SetValue(10000.5);

        string currentCultureFormat = cell.GetFormattedString();
        Assert.AreEqual("10000.5", currentCultureFormat);

        string czechCultureFormat = cell.GetFormattedString(CultureInfo.GetCultureInfo("cs-CZ"));
        Assert.AreEqual("10000,5", czechCultureFormat);
    }

    [Test]
    public void ReadAndWriteColumnNumberFormat()
    {
        using (MemoryStream memoryStream = new())
        {
            using (XLWorkbook wb = new())
            {
                IXLWorksheet ws = wb.AddWorksheet();
                IXLColumn sourceColumn = ws.Column(1);
                sourceColumn.Style.NumberFormat.Format = "0.000";
                wb.SaveAs(memoryStream);
            }

            memoryStream.Position = 0;

            using (XLWorkbook wb = new(memoryStream))
            {
                IXLColumn column = wb.Worksheets.Single().Column(1);
                Assert.AreEqual("0.000", column.Style.NumberFormat.Format);
            }
        }
    }

    [Test]
    public void XLNumberFormatGetHashCodeIsCaseSensitive()
    {
        XLNumberFormat numberFormatKey1 = new("MM");
        XLNumberFormat numberFormatKey2 = new("mm");

        Assert.AreNotEqual(numberFormatKey1.GetHashCode(), numberFormatKey2.GetHashCode());
    }

    [Test]
    public void XLNumberFormatEqualsIsCaseSensitive()
    {
        XLNumberFormat numberFormatKey1 = new("MM");
        XLNumberFormat numberFormatKey2 = new("mm");

        Assert.IsFalse(numberFormatKey1.Equals(numberFormatKey2));
    }

    [Test]
    public void AddCustomNumberFormatsToFileWithNonSequentialNumberFormatIds()
    {
        TestHelper.LoadModifyAndCompare(
            @"Other\NumberFormats\NonSequentialNumberFormatsIds-Input.xlsx",
            wb =>
            {
                IXLWorksheet ws = wb.Worksheet("Sheet1");

                string format = "\"P\" #,##0.00; \"N\" #,##0.00;0;@";
                ws.Cell(5, 1).Value = 1.2;
                ws.Cell(5, 1).Style.NumberFormat.Format = format;
                ws.Cell(5, 2).Value = -1.2;
                ws.Cell(5, 2).Style.NumberFormat.Format = format;
            },
            @"Other\NumberFormats\NonSequentialNumberFormatsIds-Output.xlsx"
        );
    }

    [Test]
    public void NumberFormatIdSetsFormatToPredefinedFormat()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLStyle cellFormat = ws.Cell("A1").Style;
        const int predefinedFormatId = (int)XLPredefinedFormat.Number.Precision2;

        cellFormat.NumberFormat.NumberFormatId = predefinedFormatId;

        Assert.AreEqual(predefinedFormatId, cellFormat.NumberFormat.NumberFormatId);
        Assert.AreEqual("0.00", cellFormat.NumberFormat.Format);
    }

    [Test]
    public void NumberFormatIdThrowsOnNonPredefinedFormats()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            ws.Cell("A1").Style.NumberFormat.NumberFormatId = 160
        );
    }

    [TestCase("0.000000 Cute", -1)]
    [TestCase("0.00", XLPredefinedFormat.Number.Precision2)]
    public void FormatSetsNumberFormat(string numberFormat, int numFmtId)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        ws.Cell("A1").Style.NumberFormat.Format = numberFormat;

        Assert.AreEqual(numberFormat, ws.Cell("A1").Style.NumberFormat.Format);
        Assert.AreEqual(numFmtId, ws.Cell("A1").Style.NumberFormat.NumberFormatId);
    }

    [Test]
    public void NumberFormatCanBeSetByAssigning()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Style.NumberFormat.Format = "0.000";

        ws.Cell("A2").Style.NumberFormat = ws.Cell("A1").Style.NumberFormat;

        Assert.AreEqual("0.000", ws.Cell("A2").Style.NumberFormat.Format);
    }

    [Test]
    public void EqualComparesFormats()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1").Style.NumberFormat.Format = "0.000";
        ws.Cell("A2").Style.NumberFormat.Format = "0.000";

        Assert.AreEqual(ws.Cell("A2").Style.NumberFormat, ws.Cell("A1").Style.NumberFormat);
    }
}

using System;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.ConditionalFormats;

public class ConditionalFormatCopyTests
{
    [Test]
    public void StylesAreCreatedDuringCopy()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        IXLConditionalFormat format = ws.Range("A1:A1").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);

        XLWorkbook wb2 = new();
        IXLWorksheet ws2 = wb2.Worksheets.Add("Sheet2");
        ws2.FirstCell().CopyFrom(ws.FirstCell());
        ClassicAssert.AreEqual(
            XLColor.Blue,
            ws2.ConditionalFormats.First().Style.Fill.BackgroundColor
        ); //Added blue style
    }

    [Test]
    public void CopyConditionalFormatSingleWorksheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        IXLConditionalFormat format = ws.Range("A1:A1").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);

        ws.Cell("A1").CopyTo("B2");

        ClassicAssert.AreEqual(1, ws.ConditionalFormats.Count());
        ClassicAssert.AreEqual("A1:A1 B2:B2", ws.ConditionalFormats.First().Ranges.ToSpaceList());
    }

    [Test]
    public void CopyConditionalFormatSameRange()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet");
        IXLConditionalFormat format = ws.Range("A1:C3").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);

        ws.Cell("A1").CopyTo("B2");

        ClassicAssert.AreEqual(1, ws.ConditionalFormats.Count());
        ClassicAssert.AreEqual("A1:C3", ws.ConditionalFormats.First().Ranges.ToSpaceList());
    }

    [Test]
    public void CopyConditionalFormatsDifferentWorksheets()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        IXLConditionalFormat format = ws1.Range("A1:A1").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");
        IXLCell otherCell = ws2.Cell("B2");

        ws1.Cell("A1").CopyTo(otherCell);

        ClassicAssert.AreEqual(1, ws1.ConditionalFormats.Count());
        ClassicAssert.AreEqual(
            "Sheet1!A1:A1",
            ws1.ConditionalFormats.First().Ranges.ToSpaceList(true)
        );

        ClassicAssert.AreEqual(1, ws2.ConditionalFormats.Count());
        ClassicAssert.AreEqual(
            "Sheet2!B2:B2",
            ws2.ConditionalFormats.First().Ranges.ToSpaceList(true)
        );
    }

    [Test]
    public void FullCopyConditionalFormatSameWorksheet()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        XLConditionalFormat format = (XLConditionalFormat)ws1.Range("A1:A1").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);

        TestDelegate action = () => format.CopyTo(ws1);

        ClassicAssert.Throws(typeof(InvalidOperationException), action);
    }

    [Test]
    public void FullCopyConditionalFormatDifferentWorksheets()
    {
        XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        XLConditionalFormat format = (XLConditionalFormat)ws1.Range("A1:C3").AddConditionalFormat();
        format
            .WhenEquals(
                "=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()
            )
            .Fill.SetBackgroundColor(XLColor.Blue);
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");

        format.CopyTo(ws2);

        ClassicAssert.AreEqual(1, ws1.ConditionalFormats.Count());
        ClassicAssert.AreEqual(
            "Sheet1!A1:C3",
            ws1.ConditionalFormats.First().Ranges.ToSpaceList(true)
        );

        ClassicAssert.AreEqual(1, ws2.ConditionalFormats.Count());
        ClassicAssert.AreEqual(
            "Sheet2!A1:C3",
            ws2.ConditionalFormats.First().Ranges.ToSpaceList(true)
        );
    }
}

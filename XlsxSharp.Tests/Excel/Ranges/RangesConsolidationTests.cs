using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

public class RangesConsolidationTests
{
    [Test]
    public void ConsolidateRangesSameWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        XLRanges ranges = new(wb);
        ranges.Add(ws.Range("A1:E3"));
        ranges.Add(ws.Range("A4:B10"));
        ranges.Add(ws.Range("E2:F12"));
        ranges.Add(ws.Range("C6:I8"));
        ranges.Add(ws.Range("G9:G9"));
        ranges.Add(ws.Range("C9:D9"));
        ranges.Add(ws.Range("H9:H9"));
        ranges.Add(ws.Range("I9:I13"));
        ranges.Add(ws.Range("C4:D5"));

        List<IXLRange> consRanges = [.. ranges.Consolidate()];

        ClassicAssert.AreEqual(6, consRanges.Count);
        ClassicAssert.AreEqual("A1:E9", consRanges[0].RangeAddress.ToString());
        ClassicAssert.AreEqual("F2:F12", consRanges[1].RangeAddress.ToString());
        ClassicAssert.AreEqual("G6:I9", consRanges[2].RangeAddress.ToString());
        ClassicAssert.AreEqual("A10:B10", consRanges[3].RangeAddress.ToString());
        ClassicAssert.AreEqual("E10:E12", consRanges[4].RangeAddress.ToString());
        ClassicAssert.AreEqual("I10:I13", consRanges[5].RangeAddress.ToString());
    }

    [Test]
    public void ConsolidateWideRangesSameWorksheet()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        XLRanges ranges = new(wb);
        ranges.Add(ws.Row(5));
        ranges.Add(ws.Row(7));
        ranges.Add(ws.Row(6));
        ranges.Add(ws.Column("D"));
        ranges.Add(ws.Column("F"));
        ranges.Add(ws.Column("E"));

        List<IXLRange> consRanges =
        [
            .. ranges
                .Consolidate()
                .OrderBy(r => r.Worksheet.Name)
                .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber),
        ];

        ClassicAssert.AreEqual(3, consRanges.Count);
        ClassicAssert.AreEqual("D:F", consRanges[0].RangeAddress.ToString());
        ClassicAssert.AreEqual("A5:C7", consRanges[1].RangeAddress.ToString());
        ClassicAssert.AreEqual("G5:XFD7", consRanges[2].RangeAddress.ToString());
    }

    [Test]
    public void ConsolidateRangesDifferentWorksheets()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws1 = wb.Worksheets.Add("Sheet1");
        IXLWorksheet ws2 = wb.Worksheets.Add("Sheet2");
        XLRanges ranges = new(wb);
        ranges.Add(ws1.Range("A1:E3"));
        ranges.Add(ws1.Range("A4:B10"));
        ranges.Add(ws1.Range("E2:F12"));
        ranges.Add(ws1.Range("C6:I8"));
        ranges.Add(ws1.Range("G9:G9"));

        ranges.Add(ws2.Row(5));
        ranges.Add(ws2.Row(7));
        ranges.Add(ws2.Row(6));
        ranges.Add(ws2.Column("D"));
        ranges.Add(ws2.Column("F"));
        ranges.Add(ws2.Column("E"));

        ranges.Add(ws1.Range("C9:D9"));
        ranges.Add(ws1.Range("H9:H9"));
        ranges.Add(ws1.Range("I9:I13"));
        ranges.Add(ws1.Range("C4:D5"));

        List<IXLRange> consRanges =
        [
            .. ranges
                .Consolidate()
                .OrderBy(r => r.Worksheet.Name)
                .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber),
        ];

        ClassicAssert.AreEqual(9, consRanges.Count);
        ClassicAssert.AreEqual(
            "Sheet1!$A$1:$E$9",
            consRanges[0].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$F$2:$F$12",
            consRanges[1].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$G$6:$I$9",
            consRanges[2].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$A$10:$B$10",
            consRanges[3].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$E$10:$E$12",
            consRanges[4].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$I$10:$I$13",
            consRanges[5].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );

        ClassicAssert.AreEqual(
            "Sheet2!$D:$F",
            consRanges[6].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet2!$A$5:$C$7",
            consRanges[7].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "Sheet2!$G$5:$XFD$7",
            consRanges[8].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    public void ConsolidateSparsedRanges()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.Worksheets.Add("Sheet1");
        XLRanges ranges = new(wb);
        ranges.Add(ws.Range("A1:C1"));
        ranges.Add(ws.Range("E1:G1"));
        ranges.Add(ws.Range("A3:C3"));
        ranges.Add(ws.Range("E3:G3"));

        List<IXLRange> consRanges = [.. ranges.Consolidate()];

        ClassicAssert.AreEqual(4, consRanges.Count);
        ClassicAssert.AreEqual("A1:C1", consRanges[0].RangeAddress.ToString());
        ClassicAssert.AreEqual("E1:G1", consRanges[1].RangeAddress.ToString());
        ClassicAssert.AreEqual("A3:C3", consRanges[2].RangeAddress.ToString());
        ClassicAssert.AreEqual("E3:G3", consRanges[3].RangeAddress.ToString());
    }
}

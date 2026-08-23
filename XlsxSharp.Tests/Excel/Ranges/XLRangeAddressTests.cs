using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Ranges;

public class XlRangeAddressTests
{
    [Test]
    public void ToStringTest()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLRangeAddress address = ws.Cell(1, 1).AsRange().RangeAddress;

        ClassicAssert.AreEqual("A1:A1", address.ToString());
        ClassicAssert.AreEqual("Sheet1!R1C1:R1C1", address.ToString(XLReferenceStyle.R1C1, true));

        ClassicAssert.AreEqual("A1:A1", address.ToStringRelative());
        ClassicAssert.AreEqual("Sheet1!A1:A1", address.ToStringRelative(true));

        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed());
        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual(
            "Sheet1!$A$1:$A$1",
            address.ToStringFixed(XLReferenceStyle.A1, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!R1C1:R1C1",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
        ClassicAssert.AreEqual(
            "Sheet1!$A$1:$A$1",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    public void ToStringTestWithSpace()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLRangeAddress address = ws.Cell(1, 1).AsRange().RangeAddress;

        ClassicAssert.AreEqual("A1:A1", address.ToString());
        ClassicAssert.AreEqual(
            "'Sheet 1'!R1C1:R1C1",
            address.ToString(XLReferenceStyle.R1C1, true)
        );

        ClassicAssert.AreEqual("A1:A1", address.ToStringRelative());
        ClassicAssert.AreEqual("'Sheet 1'!A1:A1", address.ToStringRelative(true));

        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed());
        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1:R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("$A$1:$A$1", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual(
            "'Sheet 1'!$A$1:$A$1",
            address.ToStringFixed(XLReferenceStyle.A1, true)
        );
        ClassicAssert.AreEqual(
            "'Sheet 1'!R1C1:R1C1",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
        ClassicAssert.AreEqual(
            "'Sheet 1'!$A$1:$A$1",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    [Arguments("B2:E5", "B2:E5")]
    [Arguments("E5:B2", "B2:E5")]
    [Arguments("B5:E2", "B2:E5")]
    [Arguments("B2:E$5", "B2:E$5")]
    [Arguments("B2:$E$5", "B2:$E$5")]
    [Arguments("B$2:$E$5", "B$2:$E$5")]
    [Arguments("$B$2:$E$5", "$B$2:$E$5")]
    [Arguments("B5:E$2", "B$2:E5")]
    [Arguments("$B$5:E2", "$B2:E$5")]
    [Arguments("$B$5:E$2", "$B$2:E$5")]
    [Arguments("$B$5:$E$2", "$B$2:$E$5")]
    public void RangeAddressNormalizeTest(string inputAddress, string expectedAddress)
    {
        XLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1") as XLWorksheet;
        XLRangeAddress rangeAddress = new(ws, inputAddress);

        XLRangeAddress normalizedAddress = rangeAddress.Normalize();

        ClassicAssert.AreSame(ws, rangeAddress.Worksheet);
        ClassicAssert.AreEqual(expectedAddress, normalizedAddress.ToString());
    }

    [Test]
    public void InvalidRangeAddressToStringTest()
    {
        IXLRangeAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToString());
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.Default, true));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.R1C1, true));
    }

    [Test]
    public void InvalidRangeAddressToStringFixedTest()
    {
        IXLRangeAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToStringFixed());
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual(
            "'Sheet 1'!#REF!",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "'Sheet 1'!#REF!",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
    }

    [Test]
    public void InvalidRangeAddressToStringRelativeTest()
    {
        IXLRangeAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToStringRelative());
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToStringRelative(true));
    }

    [Test]
    public void RangeAddressOnDeletedWorksheetToStringTest()
    {
        IXLRangeAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!A1:B2", address.ToString());
        ClassicAssert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!R1C1:R2C2", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("#REF!A1:B2", address.ToString(XLReferenceStyle.Default, true));
        ClassicAssert.AreEqual("#REF!R1C1:R2C2", address.ToString(XLReferenceStyle.R1C1, true));
    }

    [Test]
    public void RangeAddressOnDeletedWorksheetToStringFixedTest()
    {
        IXLRangeAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed());
        ClassicAssert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!R1C1:R2C2", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!$A$1:$B$2", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual(
            "#REF!$A$1:$B$2",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
        ClassicAssert.AreEqual(
            "#REF!R1C1:R2C2",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
    }

    [Test]
    public void RangeAddressOnDeletedWorksheetToStringRelativeTest()
    {
        IXLRangeAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!A1:B2", address.ToStringRelative());
        ClassicAssert.AreEqual("#REF!A1:B2", address.ToStringRelative(true));
    }

    [Test]
    public void InvalidRangeAddressOnDeletedWorksheetToStringTest()
    {
        IXLRangeAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!#REF!", address.ToString());
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.R1C1, true));
    }

    [Test]
    public void InvalidRangeAddressOnDeletedWorksheetToStringFixedTest()
    {
        IXLRangeAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed());
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
    }

    [Test]
    public void InvalidRangeAddressOnDeletedWorksheetToStringRelativeTest()
    {
        IXLRangeAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringRelative());
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringRelative(true));
    }

    [Test]
    public void FullSpanAddressCannotChange()
    {
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet("Sheet1");

            IXLRange wsRange = ws.AsRange();
            IXLRange row = ws.FirstRow().RowBelow(4).AsRange();
            IXLRange column = ws.FirstColumn().ColumnRight(4).AsRange();

            ClassicAssert.AreEqual($"1:{XLHelper.MaxRowNumber}", wsRange.RangeAddress.ToString());
            ClassicAssert.AreEqual("5:5", row.RangeAddress.ToString());
            ClassicAssert.AreEqual("E:E", column.RangeAddress.ToString());

            ws.Columns("Y:Z").Delete();
            ws.Rows("9:10").Delete();

            ClassicAssert.AreEqual($"1:{XLHelper.MaxRowNumber}", wsRange.RangeAddress.ToString());
            ClassicAssert.AreEqual("5:5", row.RangeAddress.ToString());
            ClassicAssert.AreEqual("E:E", column.RangeAddress.ToString());
        }
    }

    [Test]
    public void RangeAddressIsNormalized()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet();

        XLRangeAddress rangeAddress;

        rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("A1"), ws.Cell("C3")).RangeAddress;
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.Range(ws.Cell("C3"), ws.Cell("A1")).RangeAddress;
        ClassicAssert.IsFalse(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.Range("B2:B1").RangeAddress;
        ClassicAssert.IsFalse(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.Range("B2:B10").RangeAddress;
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.Range("B:B").RangeAddress;
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.Range("2:2").RangeAddress;
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);

        rangeAddress = (XLRangeAddress)ws.RangeAddress;
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);
    }

    [Test]
    public void AsRangeTests()
    {
        XLRangeAddress rangeAddress;
        rangeAddress = new XLRangeAddress(
            new XLAddress(1, 1, false, false),
            new XLAddress(5, 5, false, false)
        );

        ClassicAssert.IsTrue(rangeAddress.IsValid);
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);
        ClassicAssert.Throws<InvalidOperationException>(() => rangeAddress.AsRange());

        XLWorksheet? ws = new XLWorkbook().AddWorksheet() as XLWorksheet;
        rangeAddress = new XLRangeAddress(
            new XLAddress(ws, 1, 1, false, false),
            new XLAddress(ws, 5, 5, false, false)
        );

        ClassicAssert.IsTrue(rangeAddress.IsValid);
        ClassicAssert.IsTrue(rangeAddress.IsNormalized);
        ClassicAssert.DoesNotThrow(() => rangeAddress.AsRange());
    }

    [Test]
    public void RelativeRanges()
    {
        IXLWorksheet ws = new XLWorkbook().AddWorksheet();

        IXLRangeAddress rangeAddress;

        rangeAddress = ws.Range("D4:E4")
            .RangeAddress.Relative(
                ws.Range("A1:E4").RangeAddress,
                ws.Range("B10:F14").RangeAddress
            );
        ClassicAssert.IsTrue(rangeAddress.IsValid);
        ClassicAssert.AreEqual("E13:F13", rangeAddress.ToString());

        rangeAddress = ws.Range("D4:E4")
            .RangeAddress.Relative(
                ws.Range("B10:F14").RangeAddress,
                ws.Range("A1:E4").RangeAddress
            );
        ClassicAssert.IsFalse(rangeAddress.IsValid);
        ClassicAssert.AreEqual("#REF!", rangeAddress.ToString());

        rangeAddress = ws.Range("C3")
            .RangeAddress.Relative(ws.Range("A1:B2").RangeAddress, ws.Range("C3").RangeAddress);
        ClassicAssert.IsTrue(rangeAddress.IsValid);
        ClassicAssert.AreEqual("E5:E5", rangeAddress.ToString());

        rangeAddress = ws.Range("B2")
            .RangeAddress.Relative(ws.Range("A1").RangeAddress, ws.Range("C3").RangeAddress);
        ClassicAssert.IsTrue(rangeAddress.IsValid);
        ClassicAssert.AreEqual("D4:D4", rangeAddress.ToString());

        rangeAddress = ws.Range("A1")
            .RangeAddress.Relative(ws.Range("B2").RangeAddress, ws.Range("A1").RangeAddress);
        ClassicAssert.IsFalse(rangeAddress.IsValid);
        ClassicAssert.AreEqual("#REF!", rangeAddress.ToString());
    }

    [Test]
    public void TestSpanProperties()
    {
        XLWorksheet? ws = new XLWorkbook().AddWorksheet() as XLWorksheet;

        XLRange? range = ws.Range("B3:E5");
        IXLRangeAddress rangeAddress = range.RangeAddress as IXLRangeAddress;
        ClassicAssert.AreEqual(4, rangeAddress.ColumnSpan);
        ClassicAssert.AreEqual(3, rangeAddress.RowSpan);
        ClassicAssert.AreEqual(12, rangeAddress.NumberOfCells);

        range = ws.Range("E5:B3");
        rangeAddress = range.RangeAddress as IXLRangeAddress;
        ClassicAssert.AreEqual(4, rangeAddress.ColumnSpan);
        ClassicAssert.AreEqual(3, rangeAddress.RowSpan);
        ClassicAssert.AreEqual(12, rangeAddress.NumberOfCells);

        rangeAddress = ProduceAddressOnDeletedWorksheet();
        ClassicAssert.AreEqual(2, rangeAddress.ColumnSpan);
        ClassicAssert.AreEqual(2, rangeAddress.RowSpan);
        ClassicAssert.AreEqual(4, rangeAddress.NumberOfCells);

        rangeAddress = ProduceInvalidAddress();
        ClassicAssert.Throws<InvalidOperationException>(() => _ = rangeAddress.ColumnSpan);
        ClassicAssert.Throws<InvalidOperationException>(() => _ = rangeAddress.RowSpan);
        ClassicAssert.Throws<InvalidOperationException>(() => _ = rangeAddress.NumberOfCells);
    }

    #region Private Methods

    private static IXLRangeAddress ProduceInvalidAddress()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLRange range = ws.Range("A1:B2");

        ws.Rows(1, 5).Delete();
        return range.RangeAddress;
    }

    private static IXLRangeAddress ProduceAddressOnDeletedWorksheet()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLRangeAddress address = ws.Range("A1:B2").RangeAddress;

        ws.Delete();
        return address;
    }

    private IXLRangeAddress ProduceInvalidAddressOnDeletedWorksheet()
    {
        IXLRangeAddress address = ProduceInvalidAddress();
        address.Worksheet.Delete();
        return address;
    }

    #endregion Private Methods
}

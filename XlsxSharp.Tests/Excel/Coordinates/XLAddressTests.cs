using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

public class XlAddressTests
{
    [Test]
    public void ToStringTest()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
        IXLAddress address = ws.Cell(1, 1).Address;

        ClassicAssert.AreEqual("A1", address.ToString());
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("Sheet1!A1", address.ToString(XLReferenceStyle.Default, true));

        ClassicAssert.AreEqual("A1", address.ToStringRelative());
        ClassicAssert.AreEqual("Sheet1!A1", address.ToStringRelative(true));

        ClassicAssert.AreEqual("$A$1", address.ToStringFixed());
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("Sheet1!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("Sheet1!R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        ClassicAssert.AreEqual(
            "Sheet1!$A$1",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    public void ToStringTestWithSpace()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLAddress address = ws.Cell(1, 1).Address;

        ClassicAssert.AreEqual("A1", address.ToString());
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("'Sheet 1'!A1", address.ToString(XLReferenceStyle.Default, true));

        ClassicAssert.AreEqual("A1", address.ToStringRelative());
        ClassicAssert.AreEqual("'Sheet 1'!A1", address.ToStringRelative(true));

        ClassicAssert.AreEqual("$A$1", address.ToStringFixed());
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("'Sheet 1'!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual(
            "'Sheet 1'!R1C1",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
        ClassicAssert.AreEqual(
            "'Sheet 1'!$A$1",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    public void InvalidAddressToStringTest()
    {
        IXLAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToString());
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToString(XLReferenceStyle.Default, true));
    }

    [Test]
    public void InvalidAddressToStringFixedTest()
    {
        IXLAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToStringFixed());
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual(
            "'Sheet 1'!#REF!",
            address.ToStringFixed(XLReferenceStyle.R1C1, true)
        );
        ClassicAssert.AreEqual(
            "'Sheet 1'!#REF!",
            address.ToStringFixed(XLReferenceStyle.Default, true)
        );
    }

    [Test]
    public void InvalidAddressToStringRelativeTest()
    {
        IXLAddress address = ProduceInvalidAddress();

        ClassicAssert.AreEqual("#REF!", address.ToStringRelative());
        ClassicAssert.AreEqual("'Sheet 1'!#REF!", address.ToStringRelative(true));
    }

    [Test]
    public void AddressOnDeletedWorksheetToStringTest()
    {
        IXLAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("A1", address.ToString());
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("A1", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!A1", address.ToString(XLReferenceStyle.Default, true));
    }

    [Test]
    public void AddressOnDeletedWorksheetToStringFixedTest()
    {
        IXLAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("$A$1", address.ToStringFixed());
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("R1C1", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("$A$1", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!$A$1", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("#REF!R1C1", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        ClassicAssert.AreEqual("#REF!$A$1", address.ToStringFixed(XLReferenceStyle.Default, true));
    }

    [Test]
    public void AddressOnDeletedWorksheetToStringRelativeTest()
    {
        IXLAddress address = ProduceAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("A1", address.ToStringRelative());
        ClassicAssert.AreEqual("#REF!A1", address.ToStringRelative(true));
    }

    [Test]
    public void InvalidAddressOnDeletedWorksheetToStringTest()
    {
        IXLAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!", address.ToString());
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!", address.ToString(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToString(XLReferenceStyle.Default, true));
    }

    [Test]
    public void InvalidAddressOnDeletedWorksheetToStringFixedTest()
    {
        IXLAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!", address.ToStringFixed());
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.A1));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.R1C1));
        ClassicAssert.AreEqual("#REF!", address.ToStringFixed(XLReferenceStyle.Default));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.A1, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.R1C1, true));
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringFixed(XLReferenceStyle.Default, true));
    }

    [Test]
    public void InvalidAddressOnDeletedWorksheetToStringRelativeTest()
    {
        IXLAddress address = this.ProduceInvalidAddressOnDeletedWorksheet();

        ClassicAssert.AreEqual("#REF!", address.ToStringRelative());
        ClassicAssert.AreEqual("#REF!#REF!", address.ToStringRelative(true));
    }

    #region Private Methods

    private static IXLAddress ProduceInvalidAddress()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLRange range = ws.Range("A1:B2");

        ws.Rows(1, 5).Delete();
        return range.RangeAddress.FirstAddress;
    }

    private static IXLAddress ProduceAddressOnDeletedWorksheet()
    {
        IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet 1");
        IXLAddress address = ws.Cell("A1").Address;

        ws.Delete();
        return address;
    }

    private IXLAddress ProduceInvalidAddressOnDeletedWorksheet()
    {
        IXLAddress address = ProduceInvalidAddress();
        address.Worksheet.Delete();
        return address;
    }

    #endregion Private Methods
}

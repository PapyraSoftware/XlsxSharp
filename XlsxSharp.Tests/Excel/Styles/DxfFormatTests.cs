using System;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;

namespace XlsxSharp.Tests.Excel.Styles;

/// <summary>
/// Test of <see cref="XLDxFormat"/>.
/// </summary>
internal class DxfFormatTests
{
    [Test]
    public void Assign_dxf_to_different_dxf()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat source = ws.Range("A1").AddConditionalFormat();
        source.Style.Fill.BackgroundColor = XLColor.Red;
        IXLConditionalFormat target = ws.Range("B1").AddConditionalFormat();

        target.Style = source.Style;

        ClassicAssert.AreEqual(XLColor.Red, target.Style.Fill.BackgroundColor);

        // Copy was deep, changes to the source don't affect the copy
        source.Style.Fill.BackgroundColor = XLColor.Green;
        ClassicAssert.AreEqual(XLColor.Red, target.Style.Fill.BackgroundColor);
    }

    [Test]
    public void Cant_copy_cell_format_to_dxf()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();

        ClassicAssert.Throws<NotSupportedException>(() => cf.Style = ws.Cell("B1").Style);
    }

    [Test]
    public void IncludeQuotePrefix_always_returns_false()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();
        ClassicAssert.IsFalse(cf.Style.IncludeQuotePrefix);
    }

    [Test]
    public void IncludeQuotePrefix_cant_be_changed()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();
        ClassicAssert.Throws<NotSupportedException>(() => cf.Style.IncludeQuotePrefix = true);
    }

    [Test]
    [Arguments(-1, "$0.00")]
    [Arguments(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_format(int numFmtId, string format)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetFormat(format);

        ClassicAssert.AreEqual(format, cf.Style.NumberFormat.Format);
        ClassicAssert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
    }

    [Test]
    [Arguments(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_number_format_id(int numFmtId, string format)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetNumberFormatId(numFmtId);

        ClassicAssert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        ClassicAssert.AreEqual(format, cf.Style.NumberFormat.Format);
    }

    [Test]
    [Arguments(-1, "$0.00")]
    [Arguments(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_by_assigning_number_format(int numFmtId, string format)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLNumberFormat numberFormat = ws.Range("A1")
            .AddConditionalFormat()
            .Style.NumberFormat.SetFormat(format)
            .NumberFormat;
        IXLConditionalFormat cf = ws.Range("A2").AddConditionalFormat();

        cf.Style.NumberFormat = numberFormat;

        ClassicAssert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        ClassicAssert.AreEqual(format, cf.Style.NumberFormat.Format);
    }

    [Test]
    public void DateFormat_returns_number_format()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();
        cf.Style.NumberFormat.SetFormat("00.0");

        ClassicAssert.AreEqual("00.0", cf.Style.DateFormat.Format);
    }

    [Test]
    public void Protection_default_values_are_same_as_in_OOXML()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLProtection protection = ws.Range("A1").AddConditionalFormat().Style.Protection;

        ClassicAssert.True(protection.Locked);
        ClassicAssert.False(protection.Hidden);
    }

    [Test]
    public void Protection_locked_can_be_set()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Protection.SetLocked(false);
        ClassicAssert.False(cf.Style.Protection.Locked);
    }

    [Test]
    public void Protection_hidden_can_be_set()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.Protection.SetHidden(true);
        ClassicAssert.True(cf.Style.Protection.Hidden);
    }
}

using System;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles;

/// <summary>
/// Test of <see cref="XLDxFormat"/>.
/// </summary>
[TestFixture]
internal class DxfFormatTests
{
    [Test]
    public void Assign_dxf_to_different_dxf()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var source = ws.Range("A1").AddConditionalFormat();
        source.Style.Fill.BackgroundColor = XLColor.Red;
        var target = ws.Range("B1").AddConditionalFormat();

        target.Style = source.Style;

        Assert.That(target.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

        // Copy was deep, changes to the source don't affect the copy
        source.Style.Fill.BackgroundColor = XLColor.Green;
        Assert.That(target.Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));
    }

    [Test]
    public void Cant_copy_cell_format_to_dxf()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        Assert.That(() => cf.Style = ws.Cell("B1").Style, Throws.TypeOf<NotSupportedException>());
    }

    [Test]
    public void IncludeQuotePrefix_always_returns_false()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();
        Assert.IsFalse(cf.Style.IncludeQuotePrefix);
    }

    [Test]
    public void IncludeQuotePrefix_cant_be_changed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();
        Assert.That(() => cf.Style.IncludeQuotePrefix = true, Throws.TypeOf<NotSupportedException>());
    }

    [TestCase(-1, "$0.00")]
    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_format(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetFormat(format);

        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
    }

    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_through_number_format_id(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cf = ws.Range("A1").AddConditionalFormat();

        cf.Style.NumberFormat.SetNumberFormatId(numFmtId);

        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
    }

    [TestCase(-1, "$0.00")]
    [TestCase(XLPredefinedFormat.Number.Integer, "0")]
    public void NumberFormat_can_be_set_by_assigning_number_format(int numFmtId, string format)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var numberFormat = ws.Range("A1").AddConditionalFormat().Style.NumberFormat.SetFormat(format).NumberFormat;
        var cf = ws.Range("A2").AddConditionalFormat();

        cf.Style.NumberFormat = numberFormat;

        Assert.AreEqual(numFmtId, cf.Style.NumberFormat.NumberFormatId);
        Assert.AreEqual(format, cf.Style.NumberFormat.Format);
    }
}

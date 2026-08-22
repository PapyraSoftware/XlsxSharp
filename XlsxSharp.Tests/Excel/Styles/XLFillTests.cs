using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.Styles;

[TestFixture]
public class XLFillTests
{
    [Test]
    [TestCaseSource(nameof(FillApiSetters))]
    public void Fill_property_can_be_individually_set(FormatTestCase<IXLFill> testCase)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        // Set two-color pattern, so setting individual property doesn't trigger special logic
        IXLStyle cellFormat = ws.Cell("B2")
            .Style.Fill.SetPatternType(XLFillPatternValues.LightGrid)
            .Fill.SetBackgroundColor(XLColor.Aqua)
            .Fill.SetPatternColor(XLColor.Lemon);

        foreach (object testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellFormat.Fill, testValue);
            object setValue = testCase.GetPropertyValue(cellFormat.Fill);
            Assert.AreEqual(testValue, setValue);
        }
    }

    private static IEnumerable<FormatTestCase<IXLFill>> FillApiSetters()
    {
        XLFillPatternValues[] patternValues = EnumPolyfill.GetValues<XLFillPatternValues>();
        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.PatternType,
            (fill, value) => fill.PatternType = value,
            patternValues
        );
        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.PatternType,
            (fill, value) => fill.SetPatternType(value),
            patternValues
        );

        XLColor[] colors = [XLColor.Black, XLColor.Red, XLColor.Automatic, XLColor.Transparent];
        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.BackgroundColor,
            (fill, value) => fill.BackgroundColor = value,
            colors
        );
        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.BackgroundColor,
            (fill, value) => fill.SetBackgroundColor(value),
            colors
        );

        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.PatternColor,
            (fill, value) => fill.PatternColor = value,
            colors
        );
        yield return FormatTestCase<IXLFill>.ForFill(
            fill => fill.PatternColor,
            (fill, value) => fill.SetPatternColor(value),
            colors
        );
    }

    [Test]
    public void BackgroundColor_keeps_pattern_on_two_color_patterns()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLFill fill = ws.Cell("A1").Style.Fill;
        fill.PatternType = XLFillPatternValues.LightGrid;
        Assert.AreEqual(XLFillPatternValues.LightGrid, fill.PatternType);

        fill.BackgroundColor = XLColor.Blue;

        Assert.AreEqual(XLFillPatternValues.LightGrid, fill.PatternType);
    }

    [Test]
    public void BackgroundColor_sets_pattern_to_solid_when_original_pattern_was_none()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLFill fill = ws.Cell("A1").Style.Fill;
        Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);

        fill.BackgroundColor = XLColor.Blue;

        Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);
    }

    [Test]
    public void BackgroundColor_set_to_transparent_color_sets_pattern_to_none()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLFill fill = ws.Cell("A1").Style.Fill;
        fill.BackgroundColor = XLColor.Red;
        Assert.AreEqual(XLFillPatternValues.Solid, fill.PatternType);

        fill.BackgroundColor = XLColor.Automatic;

        Assert.AreEqual(XLFillPatternValues.None, fill.PatternType);
    }

    [Test]
    public void BackgroundPatternEqualCheck()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLFill fill1 = ws.Cell("A1").Style.Fill;
        fill1.BackgroundColor = XLColor.Blue;
        IXLFill fill2 = ws.Cell("A2").Style.Fill;
        fill2.BackgroundColor = XLColor.Blue;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
    }

    [Test]
    public void BackgroundPatternNotEqualCheck()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLFill fill1 = ws.Cell("A1").Style.Fill;
        fill1.PatternType = XLFillPatternValues.Solid;
        fill1.BackgroundColor = XLColor.Blue;

        IXLFill fill2 = ws.Cell("A2").Style.Fill;
        fill2.PatternType = XLFillPatternValues.Solid;
        fill2.BackgroundColor = XLColor.Red;

        Assert.IsFalse(fill1.Equals(fill2));
    }

    [Test]
    public void FillsWithStyleNoneAreEqualToFillWithAutomaticColor()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLFill fill1 = ws.Cell("A1")
            .Style.Fill.SetBackgroundColor(XLColor.ElectricUltramarine)
            .Fill.SetPatternType(XLFillPatternValues.None)
            .Fill;
        IXLFill fill2 = ws.Cell("A2")
            .Style.Fill.SetBackgroundColor(XLColor.EtonBlue)
            .Fill.SetPatternType(XLFillPatternValues.None)
            .Fill;
        IXLFill fill3 = ws.Cell("A3").Style.Fill.SetBackgroundColor(XLColor.Automatic).Fill;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.IsTrue(fill1.Equals(fill3));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
        Assert.AreEqual(fill1.GetHashCode(), fill3.GetHashCode());
    }

    [Test]
    public void SolidFillsWithDifferentPatternColorEqual()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();

        IXLFill fill1 = ws.Cell("A1").Style.Fill;
        fill1.PatternType = XLFillPatternValues.Solid;
        fill1.BackgroundColor = XLColor.Red;
        fill1.PatternColor = XLColor.Blue;

        IXLFill fill2 = ws.Cell("A2").Style.Fill;
        fill2.PatternType = XLFillPatternValues.Solid;
        fill2.BackgroundColor = XLColor.Red;
        fill2.PatternColor = XLColor.Green;

        Assert.IsTrue(fill1.Equals(fill2));
        Assert.AreEqual(fill1.GetHashCode(), fill2.GetHashCode());
    }

    [Test]
    public void BackgroundWithConditionalFormat()
    {
        XLWorkbook workbook = new();
        IXLWorksheet worksheet = workbook.AddWorksheet("Test");
        worksheet.Cell(2, 2).SetValue("Text");
        IXLConditionalFormat cf = worksheet.Cell(2, 2).AddConditionalFormat();
        IXLStyle style = cf.WhenNotBlank();
        style = style
            .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
            .Border.SetOutsideBorderColor(XLColor.Blue);

        Assert.AreEqual(style.Border.BottomBorder, XLBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.TopBorder, XLBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.LeftBorder, XLBorderStyleValues.Thick);
        Assert.AreEqual(style.Border.RightBorder, XLBorderStyleValues.Thick);

        Assert.AreEqual(style.Border.BottomBorderColor, XLColor.Blue);
        Assert.AreEqual(style.Border.TopBorderColor, XLColor.Blue);
        Assert.AreEqual(style.Border.LeftBorderColor, XLColor.Blue);
        Assert.AreEqual(style.Border.RightBorderColor, XLColor.Blue);
    }

    [Test]
    public void LoadAndSaveDxfBackgroundFill()
    {
        // The cells in test file have default format with a white background. Then, there are two CF:
        // * If value = 5, apply with automatic color background (=white)
        // * If value <> 5, apply red background
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\DxfBackgroundFill\inputfile.xlsx",
            @"Other\StyleReferenceFiles\DxfBackgroundFill\output.xlsx"
        );
    }

    [Test]
    public void ReservedFills_ReplaceWithPredefinedValues()
    {
        // If attribute or whole predefined fill is missing from the file, save predefined values
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-SavePredefinedValues-Output.xlsx"
        );
    }

    [Test]
    public void ReservedFills_MoveFillsFromReservedPositions()
    {
        // If the input doesn't have expected fill values at the reserved position s0 and 1 (can only happen
        // for non-excel sources, excel always has correct values), put expected fill at 0 and 1, but save original
        // fills to different positions if they are used.
        TestHelper.LoadSaveAndCompare(
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Input.xlsx",
            @"Other\StyleReferenceFiles\FillAtReservedPosition-MoveFill-Output.xlsx"
        );
    }
}

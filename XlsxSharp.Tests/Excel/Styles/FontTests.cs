using System;
using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.Styles;

public class FontTests
{
    private readonly XLFontFormatValue defaultFormat = XLFontFormatValue.Default;

    [Test]
    public void XlFontFormatValueGetHashCodeIsCaseInsensitive()
    {
        XLFontFormatValue fontKey1 = this.defaultFormat with { Name = "Arial" };
        XLFontFormatValue fontKey2 = this.defaultFormat with { Name = "Times New Roman" };
        XLFontFormatValue fontKey3 = this.defaultFormat with { Name = "TIMES NEW ROMAN" };

        Assert.AreNotEqual(fontKey1.GetHashCode(), fontKey2.GetHashCode());
        Assert.AreEqual(fontKey2.GetHashCode(), fontKey3.GetHashCode());
    }

    [Test]
    public void XlFontFormatValueEqualsIsCaseInsensitive()
    {
        XLFontFormatValue fontKey1 = this.defaultFormat with { Name = "Arial" };
        XLFontFormatValue fontKey2 = this.defaultFormat with { Name = "Times New Roman" };
        XLFontFormatValue fontKey3 = this.defaultFormat with { Name = "TIMES NEW ROMAN" };

        Assert.IsFalse(fontKey1.Equals(fontKey2));
        Assert.IsTrue(fontKey2.Equals(fontKey3));
    }

    [Test]
    [TestCaseSource(nameof(FontApiSetters))]
    public void FontPropertyCanBeIndividuallySet(FormatTestCase<IXLFont> testCase)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLCellFormat cellsFormat = ((XLCells)ws.Cells("A1:C4")).Format;
        XLCellFormat cellFormat = ((XLCell)ws.Cell("B2")).Format;

        foreach (object testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cellsFormat.Font, testValue);
            object setValue = testCase.GetPropertyValue(cellFormat.Font);
            Assert.AreEqual(testValue, setValue);
        }
    }

    [Test]
    [TestCaseSource(nameof(FontApiSetters))]
    public void DxfFontPropertyCanBeIndividuallySet(FormatTestCase<IXLFont> testCase)
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLConditionalFormat cf = ws.AddConditionalFormat();

        foreach (object testValue in testCase.Values)
        {
            testCase.SetPropertyValue(cf.Style.Font, testValue);
            object setValue = testCase.GetPropertyValue(cf.Style.Font);
            Assert.AreEqual(testValue, setValue);
        }
    }

    [Test]
    public void FontCanBeSetByAssigningFont()
    {
        // Arrange
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell("A1")
            .Style.Font.SetBold()
            .Font.SetItalic()
            .Font.SetUnderline(XLFontUnderlineValues.DoubleAccounting)
            .Font.SetStrikethrough()
            .Font.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Superscript)
            .Font.SetShadow()
            .Font.SetFontSize(25)
            .Font.SetFontColor(XLColor.Red)
            .Font.SetFontName("Arial")
            .Font.SetFontFamilyNumbering(XLFontFamilyNumberingValues.Decorative)
            .Font.SetFontCharSet(XLFontCharSet.Hangul)
            .Font.SetFontScheme(XLFontScheme.Minor);

        // Act
        ws.Cell("A2").Style.Font = ws.Cell("A1").Style.Font;

        // Assert
        IXLFont copiedFont = ws.Cell("A2").Style.Font;
        Assert.IsTrue(copiedFont.Bold);
        Assert.IsTrue(copiedFont.Italic);
        Assert.AreEqual(XLFontUnderlineValues.DoubleAccounting, copiedFont.Underline);
        Assert.IsTrue(copiedFont.Strikethrough);
        Assert.AreEqual(
            XLFontVerticalTextAlignmentValues.Superscript,
            copiedFont.VerticalAlignment
        );
        Assert.IsTrue(copiedFont.Shadow);
        Assert.AreEqual(25, copiedFont.FontSize);
        Assert.AreEqual(XLColor.Red, copiedFont.FontColor);
        Assert.AreEqual("Arial", copiedFont.FontName);
        Assert.AreEqual(XLFontFamilyNumberingValues.Decorative, copiedFont.FontFamilyNumbering);
        Assert.AreEqual(XLFontCharSet.Hangul, copiedFont.FontCharSet);
        Assert.AreEqual(XLFontScheme.Minor, copiedFont.FontScheme);
    }

    [Test]
    public void FontCanBeCheckedForEquality()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        IXLFont testFont = ws.Cell("A1").Style.Font;
        IXLFont equalFont = ws.Cell("A2").Style.Font;

        Assert.AreEqual(testFont, equalFont);
        Action<IXLFont>[] makeDifferentFont =
        [
            x => x.Bold = !x.Bold,
            x => x.Italic = !x.Italic,
            x => x.Underline = XLFontUnderlineValues.DoubleAccounting,
            x => x.Strikethrough = !x.Strikethrough,
            x => x.VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript,
            x => x.Shadow = !x.Shadow,
            x => x.FontSize = 25,
            x => x.FontColor = XLColor.Blue,
            x => x.FontName = "Arial",
            x => x.FontFamilyNumbering = XLFontFamilyNumberingValues.Decorative,
            x => x.FontCharSet = XLFontCharSet.Arabic,
            x => x.FontScheme = XLFontScheme.Minor,
        ];
        IXLCell cell = ws.Cell("A3");
        foreach (Action<IXLFont> modify in makeDifferentFont)
        {
            modify(cell.Style.Font);
            Assert.AreNotEqual(testFont, cell.Style.Font);
            cell = cell.CellRight();
        }
    }

    private static IEnumerable<object> FontApiSetters()
    {
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Bold,
            (font, value) => font.Bold = value,
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Bold,
            (font, value) => font.SetBold(value),
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Bold,
            (font, _) => font.SetBold(),
            true
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Italic,
            (font, value) => font.Italic = value,
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Italic,
            (font, value) => font.SetItalic(value),
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Italic,
            (font, _) => font.SetItalic(),
            true
        );

        XLFontUnderlineValues[] underlineValues = EnumPolyfill.GetValues<XLFontUnderlineValues>();
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Underline,
            (font, value) => font.Underline = value,
            underlineValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Underline,
            (font, value) => font.SetUnderline(value),
            underlineValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Underline,
            (font, _) => font.SetUnderline(),
            XLFontUnderlineValues.Single
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Strikethrough,
            (font, value) => font.Strikethrough = value,
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Strikethrough,
            (font, value) => font.SetStrikethrough(value),
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Strikethrough,
            (font, _) => font.SetStrikethrough(),
            true
        );

        XLFontVerticalTextAlignmentValues[] valignValues =
            EnumPolyfill.GetValues<XLFontVerticalTextAlignmentValues>();
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.VerticalAlignment,
            (font, value) => font.VerticalAlignment = value,
            valignValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.VerticalAlignment,
            (font, value) => font.SetVerticalAlignment(value),
            valignValues
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Shadow,
            (font, value) => font.Shadow = value,
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Shadow,
            (font, value) => font.SetShadow(value),
            true,
            false
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.Shadow,
            (font, _) => font.SetShadow(),
            true
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontSize,
            (font, value) => font.FontSize = value,
            1,
            15,
            409.55
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontSize,
            (font, value) => font.SetFontSize(value),
            1,
            15,
            409.55
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontColor,
            (font, value) => font.FontColor = value,
            XLColor.Black,
            XLColor.Red,
            XLColor.Automatic
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontColor,
            (font, value) => font.SetFontColor(value),
            XLColor.Black,
            XLColor.Red,
            XLColor.Automatic
        );

        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontName,
            (font, value) => font.FontName = value,
            "Calibri",
            "Arial",
            "Consolas"
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontName,
            (font, value) => font.SetFontName(value),
            "Calibri",
            "Arial",
            "Consolas"
        );

        XLFontFamilyNumberingValues[] familyValues =
            EnumPolyfill.GetValues<XLFontFamilyNumberingValues>();
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontFamilyNumbering,
            (font, value) => font.FontFamilyNumbering = value,
            familyValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontFamilyNumbering,
            (font, value) => font.SetFontFamilyNumbering(value),
            familyValues
        );

        XLFontCharSet[] charsetValues = EnumPolyfill.GetValues<XLFontCharSet>();
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontCharSet,
            (font, value) => font.FontCharSet = value,
            charsetValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontCharSet,
            (font, value) => font.SetFontCharSet(value),
            charsetValues
        );

        XLFontScheme[] schemeValues = EnumPolyfill.GetValues<XLFontScheme>();
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontScheme,
            (font, value) => font.FontScheme = value,
            schemeValues
        );
        yield return FormatTestCase<IXLFont>.ForFont(
            font => font.FontScheme,
            (font, value) => font.SetFontScheme(value),
            schemeValues
        );
    }
}

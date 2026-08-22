using System;
using System.IO;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Graphics;

namespace XlsxSharp.Tests.Graphics;

[TestFixture]
public class FontTests
{
    private readonly IXLGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;

    [TestCase]
    public void CalculatedTextWidth()
    {
        DummyFont textFont = new("Calibri", 20);
        double textWidthPt = this._engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);
        Assert.That(textWidthPt, Is.EqualTo(300));
    }

    [TestCase]
    public void CalculatedTextHeight()
    {
        DummyFont textFont = new("Calibri", 300);
        double textHeightPx = this._engine.GetTextHeight(textFont, 96);
        Assert.That(textHeightPx, Is.EqualTo(500));
    }

    [TestCase]
    public void GetMaxDigitWidth()
    {
        DummyFont textFont = new("Calibri", 11);
        double textWidthPx = this._engine.GetMaxDigitWidth(textFont, 96);
        Assert.That(textWidthPx, Is.EqualTo(7.43359375d)); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
    }

    [TestCase]
    public void DescentIsPositive()
    {
        DummyFont textFont = new("Calibri", 11);
        double textWidthPt = this._engine.GetDescent(textFont, 96);
        Assert.That(textWidthPt, Is.EqualTo(3.666666666666667d));
    }

    [TestCase]
    public void NonExistentFontUsesFallback()
    {
        DummyFont nonExistentFont = new("NonExistentFont", 100);
        DummyFont fallbackFont = new("Microsoft Sans Serif", 100);

        double nonExistentFontWidth = this._engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
        double fallbackFontWidth = this._engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
        Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

        double nonExistentFontHeight = this._engine.GetTextHeight(nonExistentFont, 96);
        double fallbackFontHeight = this._engine.GetTextHeight(fallbackFont, 96);
        Assert.That(nonExistentFontHeight, Is.EqualTo(fallbackFontHeight));
    }

    [Test]
    public void UseEmbeddedFontWhenFallbackFontIsNotPresent()
    {
        DummyFont nonExistentFont = new("SomeNonExistentFont", 11);
        DefaultGraphicEngine engine = new("NonExistentFallbackFont");
        Span<int> text = ['8'];

        GlyphBox box = engine.GetGlyphBox(text, nonExistentFont, new Dpi(96, 96));

        // Max digit width of CarlitoBare is 7, unlike MS Sans Serif which is 8
        Assert.AreEqual(7, box.AdvanceWidth);
    }

    [TestCase]
    public void CanSpecifyFallbackFontWithoutFileSystem()
    {
        using Stream fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        IXLGraphicEngine engine = DefaultGraphicEngine.CreateOnlyWithFonts(fallbackFontStream);

        DummyFont nonExistentFont = new("Nonexistent Font", 20);
        double widthOfLetterA = engine.GetTextWidth("A", nonExistentFont, 120);

        const double expectedWidthOfLetterA = 31.25d;
        Assert.AreEqual(expectedWidthOfLetterA, widthOfLetterA, 0.0001);
    }

    [TestCase]
    public void CanSpecifyExtraFontsAsStreamsWithoutFileSystem()
    {
        using Stream fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        Stream fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        IXLGraphicEngine engine = DefaultGraphicEngine.CreateOnlyWithFonts(
            fallbackFontStream,
            fontBStream
        );

        double widthOfLetterB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

        const double expectedWidthOfLetterB = 25d;
        Assert.AreEqual(expectedWidthOfLetterB, widthOfLetterB, 0.0001);
    }

    [TestCase]
    public void Issue_1916_CanMeasureSpecificArabicText()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = @"اصين";
        ws.Column(1).AdjustToContents();
    }

    private class DummyFont : IXLFontBase
    {
        public DummyFont(string name, double size)
        {
            this.FontName = name;
            this.FontSize = size;
        }

        public string FontName { get; set; }

        public double FontSize { get; set; }

        public bool Bold { get; set; }

        public bool Italic { get; set; }

        public bool Strikethrough { get; set; }

        public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;

        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

        public bool Shadow { get; set; }

        public XLColor FontColor { get; set; } = XLColor.Black;

        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; } =
            XLFontFamilyNumberingValues.NotApplicable;

        public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;

        public XLFontScheme FontScheme { get; set; }
    }
}

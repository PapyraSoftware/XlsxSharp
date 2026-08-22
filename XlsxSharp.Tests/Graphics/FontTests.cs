using System;
using System.IO;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Graphics;

namespace XlsxSharp.Tests.Graphics;

[TestFixture]
public class FontTests
{
    private readonly IXLGraphicEngine engine = DefaultGraphicEngine.Instance.Value;

    [TestCase]
    public void CalculatedTextWidth()
    {
        DummyFont textFont = new("Calibri", 20);
        double textWidthPt = this.engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);

        // SixLabors.Fonts 1.x rounded the total advance up to a whole pixel at the engine's
        // FontMetricSize of 16 (180 instead of 179.0625), which is where the previous 300 came from.
        // 3.x reports it subpixel-accurate. The value below is what the metric-compatible Carlito
        // measures; real Calibri is expected to match, but that was not verified on a Windows box.
        Assert.That(textWidthPt, Is.EqualTo(298.43747456868488d).Within(0.0001));
    }

    [TestCase]
    public void CalculatedTextHeight()
    {
        DummyFont textFont = new("Calibri", 300);
        double textHeightPx = this.engine.GetTextHeight(textFont, 96);
        Assert.That(textHeightPx, Is.EqualTo(500));
    }

    [TestCase]
    public void GetMaxDigitWidth()
    {
        DummyFont textFont = new("Calibri", 11);
        double textWidthPx = this.engine.GetMaxDigitWidth(textFont, 96);
        Assert.That(textWidthPx, Is.EqualTo(7.43359375d)); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
    }

    [TestCase]
    public void DescentIsPositive()
    {
        DummyFont textFont = new("Calibri", 11);
        double textWidthPt = this.engine.GetDescent(textFont, 96);
        Assert.That(textWidthPt, Is.EqualTo(3.666666666666667d));
    }

    [TestCase]
    public void NonExistentFontUsesFallback()
    {
        DummyFont nonExistentFont = new("NonExistentFont", 100);
        DummyFont fallbackFont = new("Microsoft Sans Serif", 100);

        double nonExistentFontWidth = this.engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
        double fallbackFontWidth = this.engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
        Assert.That(nonExistentFontWidth, Is.EqualTo(fallbackFontWidth));

        double nonExistentFontHeight = this.engine.GetTextHeight(nonExistentFont, 96);
        double fallbackFontHeight = this.engine.GetTextHeight(fallbackFont, 96);
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

        // The advance of 'A' in TestFontA is 1886/2048 em. SixLabors.Fonts 1.x rounded advances up to
        // whole pixels at the engine's FontMetricSize of 16 (giving 15/16 em); 3.x reports them
        // subpixel-accurate, so the expected value now matches the raw hmtx advance.
        const double expectedWidthOfLetterA = 30.696614583333336d;
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

        // Likewise: the advance of 'B' in TestFontB is 602/1000 em, which 1.x rounded up to 10/16 em.
        const double expectedWidthOfLetterB = 24.08d;
        Assert.AreEqual(expectedWidthOfLetterB, widthOfLetterB, 0.0001);
    }

    [TestCase]
    public void Issue1916CanMeasureSpecificArabicText()
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

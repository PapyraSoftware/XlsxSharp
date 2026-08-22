#nullable disable

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings;

namespace XlsxSharp.Graphics;

public class DefaultGraphicEngine : IXLGraphicEngine
{
    /// <summary>
    /// Carlito is a Calibri metric compatible font. This is a version stripped of everything but metric information
    /// to keep the embedded file small. It is reasonably accurate for many alphabets (contains 2531 glyphs). It has
    /// no glyph outlines, no TTF instructions, no substitutions, glyph positioning ect. It is created from Carlito
    /// font through strip-fonts.sh script.
    /// </summary>
    private const string EmbeddedFontName = "CarlitoBare";
    private const float FontMetricSize = 16f;
    private readonly ImageInfoReader[] _imageReaders =
    [
        new PngInfoReader(),
        new JpegInfoReader(),
        new GifInfoReader(),
        new TiffInfoReader(),
        new BmpInfoReader(),
        new EmfInfoReader(),
        new WmfInfoReader(),
        new WebpInfoReader(),
        new PcxInfoReader(), // Due to poor magic detection, keep last
    ];

    private readonly Lazy<IReadOnlyFontCollection> _fontCollection;
    private readonly string _fallbackFont;

    /// <summary>
    /// A font loaded font in the size <see cref="FontMetricSize"/>. There is no benefit in having multiple allocated instances, everything is just scaled at the moment.
    /// </summary>
    private readonly ConcurrentDictionary<MetricId, Font> _fonts = new();
    private readonly Func<MetricId, Font> _loadFont;

    /// <summary>
    /// Max digit width as a fraction of Em square. Multiply by font size to get pt size.
    /// </summary>
    private readonly ConcurrentDictionary<MetricId, double> _maxDigitWidths = new();
    private readonly Func<MetricId, double> _calculateMaxDigitWidth;

    /// <summary>
    /// Get a singleton instance of the engine that uses <c>Microsoft Sans Serif</c> as a fallback font.
    /// </summary>
    public static Lazy<DefaultGraphicEngine> Instance { get; } =
        new(() => new DefaultGraphicEngine("Microsoft Sans Serif"));

    /// <summary>
    /// Initialize a new instance of the engine.
    /// </summary>
    /// <param name="fallbackFont">A name of a font that is used when a font in a workbook is not available.</param>
    public DefaultGraphicEngine(string fallbackFont)
    {
        if (string.IsNullOrWhiteSpace(fallbackFont))
        {
            throw new ArgumentException(nameof(fallbackFont));
        }

        FontCollection fontCollection = new();
        AddEmbeddedFont(fontCollection);

        this._fontCollection = new Lazy<IReadOnlyFontCollection>(() =>
            fontCollection.AddSystemFonts()
        );
        this._fallbackFont = fallbackFont;
        this._loadFont = this.LoadFont;
        this._calculateMaxDigitWidth = this.CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Initialize a new instance of the engine. The engine will be able to use system fonts and fonts loaded from external sources.
    /// </summary>
    /// <remarks>Useful/necessary for environments without an access to filesystem.</remarks>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="useSystemFonts">Should engine try to use system fonts? If false, system fonts won't be loaded which can significantly speed up library startup.</param>
    /// <param name="fontStreams">Extra fonts that should be loaded to the engine.</param>
    private DefaultGraphicEngine(
        Stream fallbackFontStream,
        bool useSystemFonts,
        Stream[] fontStreams
    )
    {
        ArgumentNullException.ThrowIfNull(fallbackFontStream);

        ArgumentNullException.ThrowIfNull(fontStreams);

        FontCollection fontCollection = new();
        AddEmbeddedFont(fontCollection);
        FontFamily fallbackFamily = fontCollection.Add(fallbackFontStream);
        foreach (Stream fontStream in fontStreams)
        {
            fontCollection.Add(fontStream);
        }

        this._fontCollection = useSystemFonts
            ? new Lazy<IReadOnlyFontCollection>(() => fontCollection.AddSystemFonts())
            : new Lazy<IReadOnlyFontCollection>(() => fontCollection);
        this._fallbackFont = fallbackFamily.Name;
        this._loadFont = this.LoadFont;
        this._calculateMaxDigitWidth = this.CalculateMaxDigitWidth;
    }

    /// <summary>
    /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
    /// It ignores all system fonts and that can lead to decrease of initialization time.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Font is determined by a name and style in the worksheet, but the font name must be mapped to a font file/stream.
    /// System fonts on Windows contain hundreds of font files that have to be checked to find the correct font
    /// file for the font name and style. That means to read hundreds of files and parse data inside them.
    /// Even though SixLabors.Fonts does this only once (lazily too) and stores data in a static variable, it is
    /// an overhead that can be avoided.
    /// </para>
    /// <para>
    /// This factory method is useful in several scenarios:
    /// <list type="bullet">
    ///   <item>Client side Blazor doesn't have access to any system fonts.</item>
    ///   <item>Worksheet contains only limited number of fonts. It might be sufficient to just load few fonts we are</item>
    /// </list>
    /// </para>
    /// </remarks>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLGraphicEngine CreateOnlyWithFonts(
        Stream fallbackFontStream,
        params Stream[] fontStreams
    ) => new DefaultGraphicEngine(fallbackFontStream, false, fontStreams);

    /// <summary>
    /// Create a default graphic engine that uses only fallback font and additional fonts passed as streams.
    /// It also uses system fonts.
    /// </summary>
    /// <param name="fallbackFontStream">A stream that contains a fallback font.</param>
    /// <param name="fontStreams">Fonts that should be loaded to the engine.</param>
    public static IXLGraphicEngine CreateWithFontsAndSystemFonts(
        Stream fallbackFontStream,
        params Stream[] fontStreams
    ) => new DefaultGraphicEngine(fallbackFontStream, true, fontStreams);

    public XLPictureInfo GetPictureInfo(Stream stream, XLPictureFormat expectedFormat)
    {
        foreach (ImageInfoReader imageReader in this._imageReaders)
        {
            if (imageReader.TryGetInfo(stream, out XLPictureInfo dimensions))
            {
                return dimensions;
            }
        }

        throw new ArgumentException("Unable to determine the format of the image.");
    }

    public double GetDescent(IXLFontBase font, double dpiY)
    {
        FontMetrics metrics = this.GetMetrics(font);
        return GetDescent(font, dpiY, metrics);
    }

    private static double GetDescent(IXLFontBase font, double dpiY, FontMetrics metrics) =>
        PointsToPixels(
            -metrics.VerticalMetrics.Descender * font.FontSize / metrics.UnitsPerEm,
            dpiY
        );

    public double GetMaxDigitWidth(IXLFontBase fontBase, double dpiX)
    {
        MetricId metricId = new(fontBase);
        double maxDigitWidth = this._maxDigitWidths.GetOrAdd(
            metricId,
            this._calculateMaxDigitWidth
        );
        return PointsToPixels(maxDigitWidth * fontBase.FontSize, dpiX);
    }

    public double GetTextHeight(IXLFontBase font, double dpiY)
    {
        FontMetrics metrics = this.GetMetrics(font);
        return PointsToPixels(
            (metrics.VerticalMetrics.Ascender - 2 * metrics.VerticalMetrics.Descender)
                * font.FontSize
                / metrics.UnitsPerEm,
            dpiY
        );
    }

    public double GetTextWidth(string text, IXLFontBase fontBase, double dpiX)
    {
        Font font = this.GetFont(fontBase);
        FontRectangle dimensionsPx = TextMeasurer.MeasureAdvance(
            text,
            new TextOptions(font)
            {
                Dpi = 72, // Normalize DPI, so 1px is 1pt
                KerningMode = KerningMode.None,
            }
        );
        return PointsToPixels(dimensionsPx.Width / FontMetricSize * fontBase.FontSize, dpiX);
    }

    /// <inheritdoc />
    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
    {
        // SixLabors.Fonts don't have a way to get a glyph representation of a cluster
        // without a TextRenderer that has unacceptable performance.
        FontMetrics metric = this.GetMetrics(font);
        int advanceFu = 0;
        for (int i = 0; i < graphemeCluster.Length; ++i)
        {
            bool containsMetrics = metric.TryGetGlyphMetrics(
                new CodePoint(graphemeCluster[i]),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                null, // No palette, color fonts are not requested.
                out FontGlyphMetrics glyph
            );

            // as a fallback glyph, but it might change in the future.
            if (!containsMetrics)
            {
                continue;
            }

            advanceFu += glyph.AdvanceWidth;
        }

        double emInPx = font.FontSize / 72d * dpi.X;
        double advancePx = PointsToPixels(advanceFu * font.FontSize / metric.UnitsPerEm, dpi.X);
        double descentPx = GetDescent(font, dpi.Y, metric);
        return new GlyphBox(
            (float)Math.Round(advancePx, MidpointRounding.AwayFromZero),
            (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
            (float)Math.Round(descentPx, MidpointRounding.AwayFromZero)
        );
    }

    private FontMetrics GetMetrics(IXLFontBase fontBase)
    {
        Font font = this.GetFont(fontBase);
        return font.FontMetrics;
    }

    private Font GetFont(IXLFontBase fontBase) => this.GetFont(new MetricId(fontBase));

    private Font GetFont(MetricId metricId) => this._fonts.GetOrAdd(metricId, this._loadFont);

    private Font LoadFont(MetricId metricId)
    {
        // First try the specified fallback font. On windows, unknown fonts should use MS Sans Serif
        if (
            !this._fontCollection.Value.TryGet(metricId.Name, out FontFamily fontFamily)
            && !this._fontCollection.Value.TryGet(this._fallbackFont, out fontFamily)
        )
        {
            // If not present, e.g. it's unlikely to be present on Linux, use embedded font as an ultimate fallback.
            fontFamily = this._fontCollection.Value.Get(EmbeddedFontName);
        }

        return fontFamily.CreateFont(FontMetricSize); // Size is irrelevant for metric
    }

    private static void AddEmbeddedFont(FontCollection fontCollection)
    {
        Assembly assembly = Assembly.GetExecutingAssembly();
        const string resourcePath = "XlsxSharp.Graphics.Fonts.CarlitoBare-{0}.ttf";

        using Stream regular = assembly.GetManifestResourceStream(
            string.Format(resourcePath, "Regular")
        )!;
        fontCollection.Add(regular);

        using Stream bold = assembly.GetManifestResourceStream(
            string.Format(resourcePath, "Bold")
        )!;
        fontCollection.Add(bold);

        using Stream italic = assembly.GetManifestResourceStream(
            string.Format(resourcePath, "Italic")
        )!;
        fontCollection.Add(italic);

        using Stream boldItalic = assembly.GetManifestResourceStream(
            string.Format(resourcePath, "BoldItalic")
        )!;
        fontCollection.Add(boldItalic);
    }

    private double CalculateMaxDigitWidth(MetricId metricId)
    {
        Font font = this.GetFont(metricId);
        FontMetrics metrics = font.FontMetrics;
        int maxWidth = int.MinValue;
        for (char c = '0'; c <= '9'; ++c)
        {
            bool containsMetrics = metrics.TryGetGlyphMetrics(
                new CodePoint(c),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                null, // No palette, color fonts are not requested.
                out FontGlyphMetrics glyphMetric
            );
            if (!containsMetrics)
            {
                continue;
            }

            maxWidth = Math.Max(maxWidth, glyphMetric.AdvanceWidth);
        }
        return maxWidth / (double)metrics.UnitsPerEm;
    }

    private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

    private readonly struct MetricId : IEquatable<MetricId>
    {
        private readonly FontStyle _style;

        public MetricId(IXLFontBase fontBase)
        {
            this.Name = fontBase.FontName;
            this._style = GetFontStyle(fontBase);
        }

        public string Name { get; }

        public bool Equals(MetricId other) =>
            this.Name == other.Name && this._style == other._style;

        public override bool Equals(object obj) => obj is MetricId other && this.Equals(other);

        public override int GetHashCode() => (this.Name.GetHashCode() * 397) ^ (int)this._style;

        private static FontStyle GetFontStyle(IXLFontBase fontBase) =>
            fontBase switch
            {
                { Bold: true, Italic: true } => FontStyle.BoldItalic,
                { Bold: true } => FontStyle.Bold,
                { Italic: true } => FontStyle.Italic,
                _ => FontStyle.Regular,
            };
    }
}

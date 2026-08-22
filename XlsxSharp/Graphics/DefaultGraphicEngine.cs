#nullable disable

using System;
using System.Collections.Concurrent;
using System.Globalization;
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
    /// A system font that supplies glyphs the requested font is missing, per code point and style.
    /// A null value records that no usable font was found, so the lookup isn't repeated.
    /// </summary>
    private readonly ConcurrentDictionary<
        (int CodePoint, FontStyle Style),
        FontFamily?
    > _substitutes = new();
    private readonly Func<(int CodePoint, FontStyle Style), FontFamily?> _loadSubstitute =
        LoadSubstitute;

    /// <summary>
    /// Supplies the same substitute fonts to text shaping that <see cref="GetGlyphBox"/> uses.
    /// </summary>
    private readonly IFontFallbackResolver _fallbackResolver;

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
        this._fallbackResolver = new SubstituteFontResolver(this);
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
        this._fallbackResolver = new SubstituteFontResolver(this);
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
                // The embedded font covers Latin, Greek and Cyrillic, but not e.g. Arabic, Hebrew or
                // CJK. Those code points are resolved from system fonts, which is why a measurement
                // of such a text is not reproducible across machines the way a Latin one is.
                FontFallbackResolver = this._fallbackResolver,
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
        double advanceEm = 0;
        for (int i = 0; i < graphemeCluster.Length; ++i)
        {
            CodePoint codePoint = new(graphemeCluster[i]);

            // TryGetGlyphMetrics never returns false: a code point the font can't shape yields the
            // .notdef glyph, which has id 0 and the width of the missing-glyph box. Detecting that
            // is the only way to tell a real glyph from a missing one.
            FontMetrics glyphMetric = metric;
            if (
                !TryGetGlyph(metric, codePoint, out FontGlyphMetrics glyph)
                && TryGetSubstituteMetrics(font, codePoint, out FontMetrics substitute)
                && TryGetGlyph(substitute, codePoint, out FontGlyphMetrics substituteGlyph)
            )
            {
                // The embedded font doesn't cover this script, so a system font supplies the width.
                glyphMetric = substitute;
                glyph = substituteGlyph;
            }

            // Units per em differ between the fonts involved, so accumulate a font independent value.
            advanceEm += glyph.AdvanceWidth / (double)glyphMetric.UnitsPerEm;
        }

        double emInPx = font.FontSize / 72d * dpi.X;
        double advancePx = PointsToPixels(advanceEm * font.FontSize, dpi.X);
        double descentPx = GetDescent(font, dpi.Y, metric);
        return new GlyphBox(
            (float)Math.Round(advancePx, MidpointRounding.AwayFromZero),
            (float)Math.Round(emInPx, MidpointRounding.AwayFromZero),
            (float)Math.Round(descentPx, MidpointRounding.AwayFromZero)
        );
    }

    private static bool TryGetGlyph(
        FontMetrics metrics,
        CodePoint codePoint,
        out FontGlyphMetrics glyph
    )
    {
        // Glyph id 0 is .notdef, i.e. the font has no glyph for the code point.
        return TryGetGlyphOrNotdef(metrics, codePoint, out glyph) && glyph.GlyphId != 0;
    }

    private static bool TryGetGlyphOrNotdef(
        FontMetrics metrics,
        CodePoint codePoint,
        out FontGlyphMetrics glyph
    ) =>
        metrics.TryGetGlyphMetrics(
            codePoint,
            TextAttributes.None,
            TextDecorations.None,
            LayoutMode.HorizontalTopBottom,
            ColorFontSupport.None,
            null, // No palette, color fonts are not requested.
            out glyph
        );

    /// <summary>
    /// Find a system font that can supply a glyph the current font doesn't have. Unlike everything
    /// else in the engine the result depends on the fonts installed on the machine.
    /// </summary>
    private bool TryGetSubstituteMetrics(
        IXLFontBase font,
        CodePoint codePoint,
        out FontMetrics metrics
    )
    {
        FontFamily? family = this.GetSubstituteFamily(codePoint, MetricId.GetFontStyle(font));
        metrics = family?.CreateFont(FontMetricSize).FontMetrics;
        return metrics is not null;
    }

    private FontFamily? GetSubstituteFamily(CodePoint codePoint, FontStyle style) =>
        this._substitutes.GetOrAdd((codePoint.Value, style), this._loadSubstitute);

    private static FontFamily? LoadSubstitute((int CodePoint, FontStyle Style) key)
    {
        CodePoint codePoint = new(key.CodePoint);
        try
        {
            if (
                !SystemFonts.TryMatchCharacter(
                    codePoint,
                    key.Style,
                    null,
                    null,
                    out FontMatch match
                )
            )
            {
                return null;
            }

            FontMetrics metrics = match.Family.CreateFont(FontMetricSize).FontMetrics;

            // A font file is read lazily, on the first glyph access rather than in CreateFont, so the
            // glyph is fetched here to have both the read and the result covered by this method.
            return TryGetGlyph(metrics, codePoint, out _) ? match.Family : null;
        }
        catch (Exception e) when (e is FontException or InvalidFontFileException)
        {
            // A matched system font isn't necessarily one this library can read, e.g. macOS matches
            // CJK to a font without a 'loca' table. Fall back to the missing glyph rather than
            // letting a font on the machine break measuring a workbook.
            // Both types are caught because font loading errors derive from InvalidFontFileException,
            // which despite the name of FontException is not part of that hierarchy.
            return null;
        }
    }

    private FontMetrics GetMetrics(IXLFontBase fontBase)
    {
        Font font = this.GetFont(fontBase);
        return font.FontMetrics;
    }

    private Font GetFont(IXLFontBase fontBase) => this.GetFont(new MetricId(fontBase));

    private Font GetFont(MetricId metricId) => this._fonts.GetOrAdd(metricId, this._loadFont);

    /// <summary>
    /// Can the engine measure that font itself, or would it fall back to another one? A font name from
    /// a workbook resolves to a fallback when the machine doesn't have the font, and every measurement
    /// derived from it then belongs to the fallback instead.
    /// </summary>
    /// <param name="name">Font name, as it appears in the workbook.</param>
    internal bool IsFontAvailable(string name) =>
        SubstitutedByEmbeddedFont(name) || this._fontCollection.Value.TryGet(name, out _);

    private Font LoadFont(MetricId metricId)
    {
        // The embedded font is metric compatible with Calibri, so it is used for Calibri even when the
        // machine has the real thing installed. Calibri is the default font of a workbook, so this is
        // what makes a measurement reproducible instead of depending on what the machine happens to
        // have. It only holds for the scripts the embedded font covers; see GetTextWidth for the rest.
        if (SubstitutedByEmbeddedFont(metricId.Name))
        {
            return this._fontCollection.Value.Get(EmbeddedFontName).CreateFont(FontMetricSize);
        }

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

    /// <summary>
    /// Is the font name one the embedded font is a metric compatible stand-in for? The name comes
    /// from workbook XML, so it is compared without regard to case.
    /// </summary>
    private static bool SubstitutedByEmbeddedFont(string fontName) =>
        string.Equals(fontName, "Calibri", StringComparison.OrdinalIgnoreCase)
        || string.Equals(fontName, EmbeddedFontName, StringComparison.OrdinalIgnoreCase);

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
            // Skip digits the font has no glyph for, so the width of a missing-glyph box is not
            // mistaken for a digit width. Every column width in a workbook derives from this number.
            if (TryGetGlyph(metrics, new CodePoint(c), out FontGlyphMetrics glyphMetric))
            {
                maxWidth = Math.Max(maxWidth, glyphMetric.AdvanceWidth);
            }
        }

        if (maxWidth == int.MinValue)
        {
            // A font without any digit leaves nothing to measure. The missing-glyph box is a poor
            // width, but it is a defined one and keeps the behaviour of before this check.
            TryGetGlyphOrNotdef(metrics, new CodePoint('0'), out FontGlyphMetrics notdef);
            maxWidth = notdef.AdvanceWidth;
        }

        return maxWidth / (double)metrics.UnitsPerEm;
    }

    private static double PointsToPixels(double points, double dpi) => points / 72d * dpi;

    /// <summary>
    /// Hands text shaping the substitute fonts, so a text measured through <see cref="GetTextWidth"/>
    /// and one measured glyph by glyph agree on which font supplies a missing code point. Unusable
    /// system fonts are filtered out, which the resolver of the library itself does not do.
    /// </summary>
    private sealed class SubstituteFontResolver(DefaultGraphicEngine engine) : IFontFallbackResolver
    {
        public bool TryResolve(
            CodePoint codePoint,
            FontFamily requestedFamily,
            FontStyle style,
            CultureInfo culture,
            out FontFamily family
        )
        {
            FontFamily? substitute = engine.GetSubstituteFamily(codePoint, style);
            family = substitute ?? default;
            return substitute is not null;
        }
    }

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

        internal static FontStyle GetFontStyle(IXLFontBase fontBase) =>
            fontBase switch
            {
                { Bold: true, Italic: true } => FontStyle.BoldItalic,
                { Bold: true } => FontStyle.Bold,
                { Italic: true } => FontStyle.Italic,
                _ => FontStyle.Regular,
            };
    }
}

using System;

namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A formatting record for <see cref="XLCellFormatValue"/>.
/// </summary>
internal record XLFontFormatValue
{
    /// <summary>
    /// <para>
    /// A default font values.
    /// <list type="bullet">
    ///   <item>
    ///   When a font is loaded from the styles part, unspecified props use these values. The values
    ///   are essentially reinterpreted <em>zero</em> (other than name and size that can't be zero).
    ///   </item>
    ///   <item>
    ///   When font property has this value, it is omitted from being saved to font element (other
    ///   than font and size). Font element is pretty buggy, but these are roughly XSD default values.
    ///   </item>
    /// </list>
    /// The default values are different from normal style. Normal style is what should default
    /// format look like for user when a new workbook is created. The default font is about loading and
    /// saving.
    /// </para>
    /// <para>
    /// The font name and size are special, but the values in the property are the default fallback
    /// when even default format doesn't contain font name and size.
    /// </para>
    /// </summary>
    public static readonly XLFontFormatValue Default = new()
    {
        Name = "Calibri",
        Charset = XLFontCharSet.Ansi,
        Family = XLFontFamilyNumberingValues.NotApplicable,
        Bold = false,
        Italic = false,
        Strikethrough = false,
        Outline = false,
        Shadow = false,
        Condense = false,
        Extend = false,
        Color = XLColor.Automatic,
        Size = XLFontSize.FromPoints(11),
        Underline = XLFontUnderlineValues.None,
        VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
        Scheme = XLFontScheme.None,
    };

    public required XLFontName Name { get; init; }

    public required XLFontCharSet Charset { get; init; }

    public required XLFontFamilyNumberingValues Family { get; init; }

    public required bool Bold { get; init; }

    public required bool Italic { get; init; }

    public required bool Strikethrough { get; init; }

    public required bool Outline { get; init; }

    public required bool Shadow { get; init; }

    public required bool Condense { get; init; }

    public required bool Extend { get; init; }

    public required XLColor Color { get; init; }

    public required XLFontSize Size { get; init; }

    public required XLFontUnderlineValues Underline { get; init; }

    public required XLFontVerticalTextAlignmentValues VerticalAlignment { get; init; }

    public required XLFontScheme Scheme { get; init; }

    /// <summary>
    /// Return a registered font format equivalent to <paramref name="font"/>.
    /// </summary>
    internal static XLFontFormatValue FromFontBase(IXLFontBase font, XLWorkbookStyles styles)
    {
        XLFontFormatValue fontFormat = new()
        {
            Name = font.FontName,
            Charset = font.FontCharSet,
            Family = font.FontFamilyNumbering,
            Bold = font.Bold,
            Italic = font.Italic,
            Strikethrough = font.Strikethrough,
            Outline = Default.Outline,
            Shadow = font.Shadow,
            Condense = Default.Condense,
            Extend = Default.Extend,
            Color = font.FontColor,
            Size = XLFontSize.FromPoints(font.FontSize),
            Underline = font.Underline,
            VerticalAlignment = font.VerticalAlignment,
            Scheme = font.FontScheme,
        };
        return styles.GetRegisteredFontFormat(fontFormat, static x => x);
    }

    /// <summary>
    /// Create an adapter to font base. The adapter is not modifiable.
    /// </summary>
    internal IXLFontBase ToFontBase() => new FontBaseAdapter(this);

    private class FontBaseAdapter : IXLFontBase
    {
        private readonly XLFontFormatValue _font;

        internal FontBaseAdapter(XLFontFormatValue font) => this._font = font;

        public bool Bold
        {
            get => this._font.Bold;
            set => throw Exception();
        }

        public bool Italic
        {
            get => this._font.Italic;
            set => throw Exception();
        }
        public XLFontUnderlineValues Underline
        {
            get => this._font.Underline;
            set => throw Exception();
        }

        public bool Strikethrough
        {
            get => this._font.Strikethrough;
            set => throw Exception();
        }

        public XLFontVerticalTextAlignmentValues VerticalAlignment
        {
            get => this._font.VerticalAlignment;
            set => throw Exception();
        }

        public bool Shadow
        {
            get => this._font.Shadow;
            set => throw Exception();
        }

        public double FontSize
        {
            get => this._font.Size.Points;
            set => throw Exception();
        }

        public XLColor FontColor
        {
            get => this._font.Color;
            set => throw Exception();
        }

        public string FontName
        {
            get => this._font.Name.Text;
            set => throw Exception();
        }

        public XLFontFamilyNumberingValues FontFamilyNumbering
        {
            get => this._font.Family;
            set => throw Exception();
        }

        public XLFontCharSet FontCharSet
        {
            get => this._font.Charset;
            set => throw Exception();
        }

        public XLFontScheme FontScheme
        {
            get => this._font.Scheme;
            set => throw Exception();
        }

        private static NotSupportedException Exception() =>
            new("This is an adapter for an immutable font format.");
    }
}

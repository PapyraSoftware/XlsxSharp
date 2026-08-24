using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// API object to modify font properties of a cell format of a <see cref="IXLFormatContainer"/>.
/// </summary>
internal sealed partial class XLFontCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLFontCellFormat(XLCellFormat parent) => this._parent = parent;

    internal XLFontName Name
    {
        get => this.Resolve(static x => x.Font.Name);
        set => this.Modify(static (font, fontName) => font with { Name = fontName }, value);
    }

    internal XLFontCharSet Charset
    {
        get => this.Resolve(static x => x.Font.Charset);
        set => this.Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    internal XLFontFamilyNumberingValues Family
    {
        get => this.Resolve(static x => x.Font.Family);
        set => this.Modify(static (font, family) => font with { Family = family }, value);
    }

    internal bool Bold
    {
        get => this.Resolve(static x => x.Font.Bold);
        set => this.Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    internal bool Italic
    {
        get => this.Resolve(static x => x.Font.Italic);
        set => this.Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    internal bool Strikethrough
    {
        get => this.Resolve(static x => x.Font.Strikethrough);
        set =>
            this.Modify(
                static (font, strikethrough) => font with { Strikethrough = strikethrough },
                value
            );
    }

    internal bool Outline
    {
        get => this.Resolve(static x => x.Font.Outline);
        set => this.Modify(static (font, outline) => font with { Outline = outline }, value);
    }

    internal bool Shadow
    {
        get => this.Resolve(static x => x.Font.Shadow);
        set => this.Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    internal XLColor Color
    {
        get => this.Resolve(static x => x.Font.Color);
        set => this.Modify(static (font, color) => font with { Color = color }, value);
    }

    internal XLFontSize Size
    {
        get => this.Resolve(static x => x.Font.Size);
        set => this.Modify(static (font, size) => font with { Size = size }, value);
    }

    internal XLFontUnderlineValues Underline
    {
        get => this.Resolve(static x => x.Font.Underline);
        set => this.Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    internal XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => this.Resolve(static x => x.Font.VerticalAlignment);
        set =>
            this.Modify(
                static (font, verticalAlignment) =>
                    font with
                    {
                        VerticalAlignment = verticalAlignment,
                    },
                value
            );
    }

    internal XLFontScheme Scheme
    {
        get => this.Resolve(static x => x.Font.Scheme);
        set => this.Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    public override bool Equals(object? obj) =>
        obj is IXLFont other && (this as IEquatable<IXLFont>).Equals(other);

    public override int GetHashCode() => 0;

    private T Resolve<T>(Func<XLCellFormatValue, T> selector) => this._parent.Resolve(selector);

    private void Modify<TProperty>(
        Func<XLFontFormatValue, TProperty, XLFontFormatValue> modifyFont,
        TProperty value
    ) => this._parent.ModifyFont(modifyFont, value);

    /// <summary>
    /// A helper method to set all font properties at once (e.g, <c>someStyle.Font = otherStyle.Font</c>).
    /// </summary>
    internal void SetFont(IXLFont value) =>
        this._parent.ModifyFont(
            static (font, value) =>
                font with
                {
                    Bold = value.Bold,
                    Italic = value.Italic,
                    Underline = value.Underline,
                    Strikethrough = value.Strikethrough,
                    VerticalAlignment = value.VerticalAlignment,
                    Shadow = value.Shadow,
                    Size = XLFontSize.FromPoints(value.FontSize),
                    Color = value.FontColor,
                    Name = value.FontName,
                    Family = value.FontFamilyNumbering,
                    Charset = value.FontCharSet,
                    Scheme = value.FontScheme,
                },
            value
        );
}

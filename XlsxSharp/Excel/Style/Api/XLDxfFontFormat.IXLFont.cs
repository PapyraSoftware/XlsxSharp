using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal partial class XLDxfFontFormat : IXLFont
{
    private readonly XLFontFormatValue _defaultFont = XLFontFormatValue.Default;

    bool IXLFontBase.Bold
    {
        get => this.Resolve(static font => font.Bold, this._defaultFont.Bold);
        set => this.Modify(static (font, bold) => font with { Bold = bold }, value);
    }

    bool IXLFontBase.Italic
    {
        get => this.Resolve(static font => font.Italic, this._defaultFont.Italic);
        set => this.Modify(static (font, italic) => font with { Italic = italic }, value);
    }

    XLFontUnderlineValues IXLFontBase.Underline
    {
        get => this.Resolve(static font => font.Underline, this._defaultFont.Underline);
        set => this.Modify(static (font, underline) => font with { Underline = underline }, value);
    }

    bool IXLFontBase.Strikethrough
    {
        get => this.Resolve(static font => font.Strikethrough, this._defaultFont.Strikethrough);
        set =>
            this.Modify(
                static (font, strikethrough) => font with { Strikethrough = strikethrough },
                value
            );
    }

    XLFontVerticalTextAlignmentValues IXLFontBase.VerticalAlignment
    {
        get =>
            this.Resolve(
                static font => font.VerticalAlignment,
                this._defaultFont.VerticalAlignment
            );
        set =>
            this.Modify(static (font, vAlign) => font with { VerticalAlignment = vAlign }, value);
    }

    bool IXLFontBase.Shadow
    {
        get => this.Resolve(static font => font.Shadow, this._defaultFont.Shadow);
        set => this.Modify(static (font, shadow) => font with { Shadow = shadow }, value);
    }

    double IXLFontBase.FontSize
    {
        get => this.Resolve(static font => font.Size?.Points, this._defaultFont.Size.Points);
        set =>
            this.Modify(
                static (font, size) => font with { Size = XLFontSize.FromPoints(size) },
                value
            );
    }

    XLColor IXLFontBase.FontColor
    {
        get => this.Resolve(static font => font.Color, this._defaultFont.Color);
        set => this.Modify(static (font, color) => font with { Color = color }, value);
    }

    string IXLFontBase.FontName
    {
        get => this.Resolve(static font => font.Name?.Text, this._defaultFont.Name.Text);
        set => this.Modify(static (font, name) => font with { Name = name }, value);
    }

    XLFontFamilyNumberingValues IXLFontBase.FontFamilyNumbering
    {
        get => this.Resolve(static font => font.Family, this._defaultFont.Family);
        set => this.Modify(static (font, family) => font with { Family = family }, value);
    }

    XLFontCharSet IXLFontBase.FontCharSet
    {
        get => this.Resolve(static font => font.Charset, this._defaultFont.Charset);
        set => this.Modify(static (font, charset) => font with { Charset = charset }, value);
    }

    XLFontScheme IXLFontBase.FontScheme
    {
        get => this.Resolve(static font => font.Scheme, this._defaultFont.Scheme);
        set => this.Modify(static (font, scheme) => font with { Scheme = scheme }, value);
    }

    IXLStyle IXLFont.SetBold() => (this as IXLFont).SetBold(true);

    IXLStyle IXLFont.SetBold(bool value)
    {
        (this as IXLFont).Bold = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetItalic() => (this as IXLFont).SetItalic(true);

    IXLStyle IXLFont.SetItalic(bool value)
    {
        (this as IXLFont).Italic = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetUnderline() => (this as IXLFont).SetUnderline(XLFontUnderlineValues.Single);

    IXLStyle IXLFont.SetUnderline(XLFontUnderlineValues value)
    {
        (this as IXLFont).Underline = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetStrikethrough() => (this as IXLFont).SetStrikethrough(true);

    IXLStyle IXLFont.SetStrikethrough(bool value)
    {
        (this as IXLFont).Strikethrough = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        (this as IXLFont).VerticalAlignment = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetShadow() => (this as IXLFont).SetShadow(true);

    IXLStyle IXLFont.SetShadow(bool value)
    {
        (this as IXLFont).Shadow = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontSize(double value)
    {
        (this as IXLFont).FontSize = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontColor(XLColor value)
    {
        (this as IXLFont).FontColor = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontName(string value)
    {
        (this as IXLFont).FontName = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        (this as IXLFont).FontFamilyNumbering = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontCharSet(XLFontCharSet value)
    {
        (this as IXLFont).FontCharSet = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontScheme(XLFontScheme value)
    {
        (this as IXLFont).FontScheme = value;
        return this._parent;
    }

    bool IEquatable<IXLFont>.Equals(IXLFont? other) => throw new NotSupportedException();
}

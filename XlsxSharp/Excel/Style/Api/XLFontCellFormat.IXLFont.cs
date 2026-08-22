using System;

namespace XlsxSharp.Excel;

internal sealed partial class XLFontCellFormat : IXLFont
{
    bool IXLFontBase.Bold
    {
        get => this.Bold;
        set => this.Bold = value;
    }

    bool IXLFontBase.Italic
    {
        get => this.Italic;
        set => this.Italic = value;
    }

    XLFontUnderlineValues IXLFontBase.Underline
    {
        get => this.Underline;
        set => this.Underline = value;
    }

    bool IXLFontBase.Strikethrough
    {
        get => this.Strikethrough;
        set => this.Strikethrough = value;
    }

    XLFontVerticalTextAlignmentValues IXLFontBase.VerticalAlignment
    {
        get => this.VerticalAlignment;
        set => this.VerticalAlignment = value;
    }

    bool IXLFontBase.Shadow
    {
        get => this.Shadow;
        set => this.Shadow = value;
    }

    double IXLFontBase.FontSize
    {
        get => this.Size.Points;
        set => this.Size = XLFontSize.FromPoints(value);
    }

    XLColor IXLFontBase.FontColor
    {
        get => this.Color;
        set => this.Color = value;
    }

    string IXLFontBase.FontName
    {
        get => this.Name.Text;
        set => this.Name = value;
    }

    XLFontFamilyNumberingValues IXLFontBase.FontFamilyNumbering
    {
        get => this.Family;
        set => this.Family = value;
    }

    XLFontCharSet IXLFontBase.FontCharSet
    {
        get => this.Charset;
        set => this.Charset = value;
    }

    XLFontScheme IXLFontBase.FontScheme
    {
        get => this.Scheme;
        set => this.Scheme = value;
    }

    bool IEquatable<IXLFont>.Equals(IXLFont? other)
    {
        if (other is null)
        {
            return false;
        }

        if (this.Bold != other.Bold)
        {
            return false;
        }

        if (this.Italic != other.Italic)
        {
            return false;
        }

        if (this.Underline != other.Underline)
        {
            return false;
        }

        if (this.Strikethrough != other.Strikethrough)
        {
            return false;
        }

        if (this.VerticalAlignment != other.VerticalAlignment)
        {
            return false;
        }

        if (this.Shadow != other.Shadow)
        {
            return false;
        }

        if (!this.Size.Points.Equals(other.FontSize))
        {
            return false;
        }

        if (this.Color != other.FontColor)
        {
            return false;
        }

        if (this.Name != other.FontName)
        {
            return false;
        }

        if (this.Family != other.FontFamilyNumbering)
        {
            return false;
        }

        if (this.Charset != other.FontCharSet)
        {
            return false;
        }

        if (this.Scheme != other.FontScheme)
        {
            return false;
        }

        return true;
    }

    IXLStyle IXLFont.SetBold()
    {
        return (this as IXLFont).SetBold(true);
    }

    IXLStyle IXLFont.SetBold(bool value)
    {
        this.Bold = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetItalic()
    {
        return (this as IXLFont).SetItalic(true);
    }

    IXLStyle IXLFont.SetItalic(bool value)
    {
        this.Italic = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetUnderline()
    {
        return (this as IXLFont).SetUnderline(XLFontUnderlineValues.Single);
    }

    IXLStyle IXLFont.SetUnderline(XLFontUnderlineValues value)
    {
        this.Underline = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetStrikethrough()
    {
        return (this as IXLFont).SetStrikethrough(true);
    }

    IXLStyle IXLFont.SetStrikethrough(bool value)
    {
        this.Strikethrough = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        this.VerticalAlignment = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetShadow()
    {
        return (this as IXLFont).SetShadow(true);
    }

    IXLStyle IXLFont.SetShadow(bool value)
    {
        this.Shadow = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontSize(double value)
    {
        (this as IXLFont).FontSize = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontColor(XLColor value)
    {
        this.Color = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontName(string value)
    {
        this.Name = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        this.Family = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontCharSet(XLFontCharSet value)
    {
        this.Charset = value;
        return this._parent;
    }

    IXLStyle IXLFont.SetFontScheme(XLFontScheme value)
    {
        this.Scheme = value;
        return this._parent;
    }
}

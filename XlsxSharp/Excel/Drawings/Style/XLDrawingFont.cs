#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingFont : IXLDrawingFont
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingFont(IXLDrawingStyle style)
    {
        this._style = style;
        this.FontName = "Tahoma";
        this.FontSize = 9;
        this.Underline = XLFontUnderlineValues.None;
        this.FontColor = XLColor.FromIndex(64);
    }

    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public XLFontUnderlineValues Underline { get; set; }
    public bool Strikethrough { get; set; }
    public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
    public bool Shadow { get; set; }
    public double FontSize { get; set; }
    public XLColor FontColor { get; set; }
    public string FontName { get; set; }
    public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }
    public XLFontCharSet FontCharSet { get; set; }
    public XLFontScheme FontScheme { get; set; }

    public IXLDrawingStyle SetBold()
    {
        this.Bold = true;
        return this._style;
    }

    public IXLDrawingStyle SetBold(bool value)
    {
        this.Bold = value;
        return this._style;
    }

    public IXLDrawingStyle SetItalic()
    {
        this.Italic = true;
        return this._style;
    }

    public IXLDrawingStyle SetItalic(bool value)
    {
        this.Italic = value;
        return this._style;
    }

    public IXLDrawingStyle SetUnderline()
    {
        this.Underline = XLFontUnderlineValues.Single;
        return this._style;
    }

    public IXLDrawingStyle SetUnderline(XLFontUnderlineValues value)
    {
        this.Underline = value;
        return this._style;
    }

    public IXLDrawingStyle SetStrikethrough()
    {
        this.Strikethrough = true;
        return this._style;
    }

    public IXLDrawingStyle SetStrikethrough(bool value)
    {
        this.Strikethrough = value;
        return this._style;
    }

    public IXLDrawingStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        this.VerticalAlignment = value;
        return this._style;
    }

    public IXLDrawingStyle SetShadow()
    {
        this.Shadow = true;
        return this._style;
    }

    public IXLDrawingStyle SetShadow(bool value)
    {
        this.Shadow = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontSize(double value)
    {
        this.FontSize = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontColor(XLColor value)
    {
        this.FontColor = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontName(string value)
    {
        this.FontName = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        this.FontFamilyNumbering = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontCharSet(XLFontCharSet value)
    {
        this.FontCharSet = value;
        return this._style;
    }

    public IXLDrawingStyle SetFontScheme(XLFontScheme value)
    {
        this.FontScheme = value;
        return this._style;
    }
}

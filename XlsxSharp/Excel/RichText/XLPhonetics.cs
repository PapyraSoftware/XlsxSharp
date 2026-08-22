using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.RichText;

internal class XLPhonetics : IXLPhonetics
{
    private readonly List<IXLPhonetic> _phonetics = [];
    private readonly XLWorkbookStyles _styles;
    private readonly XLFontFormatValue _defaultFont;
    private readonly Action _onChange;
    private XLFontFormatValue _font;
    private XLPhoneticAlignment _alignment;
    private XLPhoneticType _type;

    public XLPhonetics(
        XLFontFormatValue font,
        XLFontFormatValue defaultFont,
        XLWorkbookStyles styles,
        Action onChange
    )
    {
        this._styles = styles;
        this._defaultFont = defaultFont;
        this._font = font;
        this._type = XLPhoneticType.FullWidthKatakana;
        this._alignment = XLPhoneticAlignment.Left;
        this._onChange = onChange;
    }

    public int Count => this._phonetics.Count;

    public bool Bold
    {
        get => this._font.Bold;
        set => this.ChangeFont(f => f with { Bold = value });
    }

    public bool Italic
    {
        get => this._font.Italic;
        set => this.ChangeFont(f => f with { Italic = value });
    }

    public XLFontUnderlineValues Underline
    {
        get => this._font.Underline;
        set => this.ChangeFont(f => f with { Underline = value });
    }

    public bool Strikethrough
    {
        get => this._font.Strikethrough;
        set => this.ChangeFont(f => f with { Strikethrough = value });
    }

    public XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => this._font.VerticalAlignment;
        set => this.ChangeFont(f => f with { VerticalAlignment = value });
    }

    public bool Shadow
    {
        get => this._font.Shadow;
        set => this.ChangeFont(f => f with { Shadow = value });
    }

    public double FontSize
    {
        get => this._font.Size.Points;
        set => this.ChangeFont(f => f with { Size = XLFontSize.FromPoints(value) });
    }

    public XLColor FontColor
    {
        get => this._font.Color;
        set => this.ChangeFont(f => f with { Color = value });
    }

    public string FontName
    {
        get => this._font.Name.Text;
        set => this.ChangeFont(f => f with { Name = value });
    }

    public XLFontFamilyNumberingValues FontFamilyNumbering
    {
        get => this._font.Family;
        set => this.ChangeFont(f => f with { Family = value });
    }

    public XLFontCharSet FontCharSet
    {
        get => this._font.Charset;
        set => this.ChangeFont(f => f with { Charset = value });
    }

    public XLFontScheme FontScheme
    {
        get => this._font.Scheme;
        set => this.ChangeFont(f => f with { Scheme = value });
    }

    public XLPhoneticAlignment Alignment
    {
        get => this._alignment;
        set
        {
            this._alignment = value;
            this._onChange();
        }
    }

    public XLPhoneticType Type
    {
        get => this._type;
        set
        {
            this._type = value;
            this._onChange();
        }
    }

    internal XLFontFormatValue Font => this._font;

    public IXLPhonetics SetBold()
    {
        this.Bold = true;
        return this;
    }

    public IXLPhonetics SetBold(bool value)
    {
        this.Bold = value;
        return this;
    }

    public IXLPhonetics SetItalic()
    {
        this.Italic = true;
        return this;
    }

    public IXLPhonetics SetItalic(bool value)
    {
        this.Italic = value;
        return this;
    }

    public IXLPhonetics SetUnderline()
    {
        this.Underline = XLFontUnderlineValues.Single;
        return this;
    }

    public IXLPhonetics SetUnderline(XLFontUnderlineValues value)
    {
        this.Underline = value;
        return this;
    }

    public IXLPhonetics SetStrikethrough()
    {
        this.Strikethrough = true;
        return this;
    }

    public IXLPhonetics SetStrikethrough(bool value)
    {
        this.Strikethrough = value;
        return this;
    }

    public IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        this.VerticalAlignment = value;
        return this;
    }

    public IXLPhonetics SetShadow()
    {
        this.Shadow = true;
        return this;
    }

    public IXLPhonetics SetShadow(bool value)
    {
        this.Shadow = value;
        return this;
    }

    public IXLPhonetics SetFontSize(double value)
    {
        this.FontSize = value;
        return this;
    }

    public IXLPhonetics SetFontColor(XLColor value)
    {
        this.FontColor = value;
        return this;
    }

    public IXLPhonetics SetFontName(string value)
    {
        this.FontName = value;
        return this;
    }

    public IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        this.FontFamilyNumbering = value;
        return this;
    }

    public IXLPhonetics SetFontCharSet(XLFontCharSet value)
    {
        this.FontCharSet = value;
        return this;
    }

    public IXLPhonetics SetFontScheme(XLFontScheme value)
    {
        this.FontScheme = value;
        return this;
    }

    public IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment)
    {
        this.Alignment = phoneticAlignment;
        return this;
    }

    public IXLPhonetics SetType(XLPhoneticType phoneticType)
    {
        this.Type = phoneticType;
        return this;
    }

    public IXLPhonetics Add(string text, int start, int end)
    {
        this._phonetics.Add(new XLPhonetic(text, start, end));
        this._onChange();
        return this;
    }

    public IXLPhonetics ClearText()
    {
        this._phonetics.Clear();
        this._onChange();
        return this;
    }

    public IXLPhonetics ClearFont()
    {
        this._font = this._defaultFont;
        this._onChange();
        return this;
    }

    public IEnumerator<IXLPhonetic> GetEnumerator() => this._phonetics.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    public bool Equals(IXLPhonetics? other) => this.Equals(other as XLPhonetics);

    public bool Equals(XLPhonetics? other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        if (!this._phonetics.SequenceEqual(other._phonetics))
        {
            return false;
        }

        return this._font.Equals(other._font)
            && this.Type == other.Type
            && this.Alignment == other.Alignment;
    }

    private void ChangeFont(Func<XLFontFormatValue, XLFontFormatValue> modifyFont)
    {
        this._font = this._styles.GetRegisteredFontFormat(this._font, modifyFont);
        this._onChange();
    }
}

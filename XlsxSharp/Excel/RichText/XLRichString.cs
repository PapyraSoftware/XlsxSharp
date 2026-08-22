using System;
using System.Diagnostics;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.RichText;

/// <summary>
/// An API object to modify a rich string.
/// </summary>
[DebuggerDisplay("{Text}")]
internal class XLRichString : IXLRichString
{
    private readonly XLWorkbookStyles _styles;
    private readonly IXLWithRichString _withRichString;
    private readonly Action _onChange;
    private string _text;
    private XLFontFormatValue _font;

    internal XLRichString(
        string text,
        XLFontFormatValue font,
        IXLWithRichString withRichString,
        XLWorkbookStyles styles,
        Action? onChange
    )
    {
        this._text = text;
        this._font = font;
        this._withRichString = withRichString;
        this._styles = styles;
        this._onChange = onChange ?? (() => { });
    }

    public string Text
    {
        get => this._text;
        set
        {
            this._text = value;
            this._onChange();
        }
    }

    public IXLRichString AddText(string text) => this._withRichString.AddText(text);

    public IXLRichString AddNewLine() => this.AddText(Environment.NewLine);

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

    internal XLFontFormatValue Font => this._font;

    public IXLRichString SetBold()
    {
        this.Bold = true;
        return this;
    }

    public IXLRichString SetBold(bool value)
    {
        this.Bold = value;
        return this;
    }

    public IXLRichString SetItalic()
    {
        this.Italic = true;
        return this;
    }

    public IXLRichString SetItalic(bool value)
    {
        this.Italic = value;
        return this;
    }

    public IXLRichString SetUnderline()
    {
        this.Underline = XLFontUnderlineValues.Single;
        return this;
    }

    public IXLRichString SetUnderline(XLFontUnderlineValues value)
    {
        this.Underline = value;
        return this;
    }

    public IXLRichString SetStrikethrough()
    {
        this.Strikethrough = true;
        return this;
    }

    public IXLRichString SetStrikethrough(bool value)
    {
        this.Strikethrough = value;
        return this;
    }

    public IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        this.VerticalAlignment = value;
        return this;
    }

    public IXLRichString SetShadow()
    {
        this.Shadow = true;
        return this;
    }

    public IXLRichString SetShadow(bool value)
    {
        this.Shadow = value;
        return this;
    }

    public IXLRichString SetFontSize(double value)
    {
        this.FontSize = value;
        return this;
    }

    public IXLRichString SetFontColor(XLColor value)
    {
        this.FontColor = value;
        return this;
    }

    public IXLRichString SetFontName(string value)
    {
        this.FontName = value;
        return this;
    }

    public IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        this.FontFamilyNumbering = value;
        return this;
    }

    public IXLRichString SetFontCharSet(XLFontCharSet value)
    {
        this.FontCharSet = value;
        return this;
    }

    public IXLRichString SetFontScheme(XLFontScheme value)
    {
        this.FontScheme = value;
        return this;
    }

    public override bool Equals(object? obj) => this.Equals(obj as XLRichString);

    public bool Equals(IXLRichString? other) => this.Equals(other as XLRichString);

    public bool Equals(XLRichString? other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return this.Text == other.Text && this._font.Equals(other._font);
    }

    public override int GetHashCode() =>
        // Since all properties of type are mutable, can't have different hashcode for any instance.
        // Don't ever use this class in a dictionary, e.g. SST.
        4; // Chosen by fair dice roll. Guaranteed to be random.

    private void ChangeFont(Func<XLFontFormatValue, XLFontFormatValue> modifyFont)
    {
        this._font = this._styles.GetRegisteredFontFormat(this._font, modifyFont);
        this._onChange();
    }
}

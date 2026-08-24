#nullable disable

using System.Diagnostics;
using System.Text;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.RichText;

internal class XLFormattedText<T> : IXLFormattedText<T>
{
    // TODO: Move from ancestor to children, not needed here
    protected T Container;
    protected readonly XLWorkbookStyles Styles;

    /// <summary>
    /// Font used for a new rich text run, never modified. It is generally provided by a container of the formatted text.
    /// </summary>
    private readonly XLFontFormatValue _defaultFont;
    private readonly List<XLRichString> _richTexts = [];
    private XLPhonetics _phonetics;

    protected XLFormattedText(XLFontFormatValue defaultFont, XLWorkbookStyles styles)
    {
        Debug.Assert(styles.Fonts.ContainsValue(defaultFont));
        this._defaultFont = defaultFont;
        this.Styles = styles;
    }

    IXLPhonetics IXLFormattedText<T>.Phonetics => this.Phonetics;

    public int Count => this._richTexts.Count;

    public int Length { get; private set; }

    public string Text => this.ToString();

    public bool HasPhonetics => this._phonetics is not null;

    /// <inheritdoc cref="IXLFormattedText{T}.Phonetics"/>
    internal XLPhonetics Phonetics
    {
        get =>
            this._phonetics ??= new XLPhonetics(
                this._defaultFont,
                this._defaultFont,
                this.Styles,
                this.OnContentChanged
            );
        init => this._phonetics = value;
    }

    public IXLRichString AddText(string text)
    {
        XLRichString richText = new(
            text,
            this._defaultFont,
            this,
            this.Styles,
            this.OnContentChanged
        );
        this.AddText(richText);
        this.OnContentChanged();
        return richText;
    }

    public IXLRichString AddText(string text, IXLFontBase font)
    {
        XLFontFormatValue richFont = XLFontFormatValue.FromFontBase(font, this.Styles);
        XLRichString richText = new(text, richFont, this, this.Styles, this.OnContentChanged);
        this.AddText(richText);
        this.OnContentChanged();
        return richText;
    }

    public IXLRichString AddNewLine() => this.AddText(Environment.NewLine);

    public IXLFormattedText<T> ClearText()
    {
        this.ClearContent();
        this.OnContentChanged();
        return this;
    }

    public IXLFormattedText<T> ClearFont()
    {
        string text = this.Text;
        this.ClearContent();
        this.AddText(text);
        return this;
    }

    public override string ToString()
    {
        StringBuilder sb = new(this._richTexts.Count);
        this._richTexts.ForEach(rt => sb.Append(rt.Text));
        return sb.ToString();
    }

    public IXLFormattedText<T> Substring(int index) => this.Substring(index, this.Length - index);

    public IXLFormattedText<T> Substring(int index, int length)
    {
        if (index + 1 > this.Length || (this.Length - index + 1) < length || length <= 0)
        {
            throw new IndexOutOfRangeException(
                "Index and length must refer to a location within the string."
            );
        }

        List<XLRichString> newRichTexts = [];
        XLFormattedText<T> retVal = new(this._defaultFont, this.Styles);

        int lastPosition = 0;
        foreach (XLRichString rt in this._richTexts)
        {
            if (lastPosition >= index + 1 + length) // We already have what we need
            {
                newRichTexts.Add(rt);
            }
            else if (lastPosition + rt.Text.Length >= index + 1) // Eureka!
            {
                int startIndex = index - lastPosition;

                if (startIndex > 0)
                {
                    newRichTexts.Add(
                        new XLRichString(
                            rt.Text[..startIndex],
                            rt.Font,
                            this,
                            this.Styles,
                            this.OnContentChanged
                        )
                    );
                }
                else if (startIndex < 0)
                {
                    startIndex = 0;
                }

                int leftToTake = length - retVal.Length;
                if (leftToTake > rt.Text.Length - startIndex)
                {
                    leftToTake = rt.Text.Length - startIndex;
                }

                XLRichString newRt = new(
                    rt.Text.Substring(startIndex, leftToTake),
                    rt.Font,
                    this,
                    this.Styles,
                    this.OnContentChanged
                );
                newRichTexts.Add(newRt);
                retVal.AddText(newRt);

                if (startIndex + leftToTake < rt.Text.Length)
                {
                    newRichTexts.Add(
                        new XLRichString(
                            rt.Text.Substring(startIndex + leftToTake),
                            rt.Font,
                            this,
                            this.Styles,
                            this.OnContentChanged
                        )
                    );
                }
            }
            else // We haven't reached the desired position yet
            {
                newRichTexts.Add(rt);
            }
            lastPosition += rt.Text.Length;
        }

        this._richTexts.Clear();
        this._richTexts.AddRange(newRichTexts);
        this.OnContentChanged();
        return retVal;
    }

    public IXLFormattedText<T> CopyFrom(IXLFormattedText<T> original)
    {
        this.ClearContent();
        foreach (IXLRichString richText in original)
        {
            XLFontFormatValue copyFont = XLFontFormatValue.FromFontBase(richText, this.Styles);
            XLRichString copyText = new(
                richText.Text,
                copyFont,
                this,
                this.Styles,
                this.OnContentChanged
            );
            this.AddText(copyText);
        }

        this.OnContentChanged();
        return this;
    }

    public List<XLRichString>.Enumerator GetEnumerator() => this._richTexts.GetEnumerator();

    IEnumerator<IXLRichString> IEnumerable<IXLRichString>.GetEnumerator() => this.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    public bool Bold
    {
        set => this._richTexts.ForEach(rt => rt.Bold = value);
    }
    public bool Italic
    {
        set => this._richTexts.ForEach(rt => rt.Italic = value);
    }
    public XLFontUnderlineValues Underline
    {
        set => this._richTexts.ForEach(rt => rt.Underline = value);
    }
    public bool Strikethrough
    {
        set => this._richTexts.ForEach(rt => rt.Strikethrough = value);
    }
    public XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        set => this._richTexts.ForEach(rt => rt.VerticalAlignment = value);
    }
    public bool Shadow
    {
        set => this._richTexts.ForEach(rt => rt.Shadow = value);
    }
    public double FontSize
    {
        set => this._richTexts.ForEach(rt => rt.FontSize = value);
    }
    public XLColor FontColor
    {
        set => this._richTexts.ForEach(rt => rt.FontColor = value);
    }
    public string FontName
    {
        set => this._richTexts.ForEach(rt => rt.FontName = value);
    }
    public XLFontFamilyNumberingValues FontFamilyNumbering
    {
        set => this._richTexts.ForEach(rt => rt.FontFamilyNumbering = value);
    }

    public IXLFormattedText<T> SetBold()
    {
        this.Bold = true;
        return this;
    }

    public IXLFormattedText<T> SetBold(bool value)
    {
        this.Bold = value;
        return this;
    }

    public IXLFormattedText<T> SetItalic()
    {
        this.Italic = true;
        return this;
    }

    public IXLFormattedText<T> SetItalic(bool value)
    {
        this.Italic = value;
        return this;
    }

    public IXLFormattedText<T> SetUnderline()
    {
        this.Underline = XLFontUnderlineValues.Single;
        return this;
    }

    public IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value)
    {
        this.Underline = value;
        return this;
    }

    public IXLFormattedText<T> SetStrikethrough()
    {
        this.Strikethrough = true;
        return this;
    }

    public IXLFormattedText<T> SetStrikethrough(bool value)
    {
        this.Strikethrough = value;
        return this;
    }

    public IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value)
    {
        this.VerticalAlignment = value;
        return this;
    }

    public IXLFormattedText<T> SetShadow()
    {
        this.Shadow = true;
        return this;
    }

    public IXLFormattedText<T> SetShadow(bool value)
    {
        this.Shadow = value;
        return this;
    }

    public IXLFormattedText<T> SetFontSize(double value)
    {
        this.FontSize = value;
        return this;
    }

    public IXLFormattedText<T> SetFontColor(XLColor value)
    {
        this.FontColor = value;
        return this;
    }

    public IXLFormattedText<T> SetFontName(string value)
    {
        this.FontName = value;
        return this;
    }

    public IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value)
    {
        this.FontFamilyNumbering = value;
        return this;
    }

    public bool Equals(IXLFormattedText<T> other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        if (this.Count != other.Count)
        {
            return false;
        }

        if (!this._richTexts.SequenceEqual(other))
        {
            return false;
        }

        return (this._phonetics is null && !other.HasPhonetics)
            || this.Phonetics.Equals(other.Phonetics);
    }

    protected void AddText(XLRichString richText)
    {
        this._richTexts.Add(richText);
        this.Length += richText.Text.Length;
    }

    /// <summary>
    /// This method is called every time the formatted text is changed (new runs, font props, phonetics...).
    /// </summary>
    protected virtual void OnContentChanged()
    {
        // Do nothing, intended to be overriden.
    }

    private void ClearContent()
    {
        this._richTexts.Clear();
        this.Length = 0;
    }
}

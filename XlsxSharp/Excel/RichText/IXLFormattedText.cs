#nullable disable

namespace XlsxSharp.Excel.RichText;

public interface IXLFormattedText<T>
    : IEnumerable<IXLRichString>,
        IEquatable<IXLFormattedText<T>>,
        IXLWithRichString
{
    public bool Bold { set; }
    public bool Italic { set; }
    public XLFontUnderlineValues Underline { set; }
    public bool Strikethrough { set; }
    public XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
    public bool Shadow { set; }
    public double FontSize { set; }
    public XLColor FontColor { set; }
    public string FontName { set; }
    public XLFontFamilyNumberingValues FontFamilyNumbering { set; }

    public IXLFormattedText<T> SetBold();
    public IXLFormattedText<T> SetBold(bool value);
    public IXLFormattedText<T> SetItalic();
    public IXLFormattedText<T> SetItalic(bool value);
    public IXLFormattedText<T> SetUnderline();
    public IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value);
    public IXLFormattedText<T> SetStrikethrough();
    public IXLFormattedText<T> SetStrikethrough(bool value);
    public IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
    public IXLFormattedText<T> SetShadow();
    public IXLFormattedText<T> SetShadow(bool value);
    public IXLFormattedText<T> SetFontSize(double value);
    public IXLFormattedText<T> SetFontColor(XLColor value);
    public IXLFormattedText<T> SetFontName(string value);
    public IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

    public IXLRichString AddText(string text, IXLFontBase font);
    public IXLFormattedText<T> ClearText();
    public IXLFormattedText<T> ClearFont();

    public IXLFormattedText<T> Substring(int index);
    public IXLFormattedText<T> Substring(int index, int length);

    /// <summary>
    /// Replace the text and formatting of this text by texts and formatting from the <paramref name="original"/> text.
    /// </summary>
    /// <param name="original">Original to copy from.</param>
    /// <returns>This text.</returns>
    public IXLFormattedText<T> CopyFrom(IXLFormattedText<T> original);

    /// <summary>
    /// How many rich strings is the formatted text composed of.
    /// </summary>
    public int Count { get; }

    /// <summary>
    /// Length of the whole formatted text.
    /// </summary>
    public int Length { get; }

    /// <summary>
    /// Get text of the whole formatted text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Does this text has phonetics? Unlike accessing the <see cref="Phonetics"/> property, this method
    /// doesn't create a new instance on access.
    /// </summary>
    public bool HasPhonetics { get; }

    /// <summary>
    /// Get or create phonetics for the text. Use <see cref="HasPhonetics"/> to check for existence to avoid unnecessary creation.
    /// </summary>
    public IXLPhonetics Phonetics { get; }
}

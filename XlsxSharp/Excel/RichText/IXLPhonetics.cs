using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.RichText;

public enum XLPhoneticAlignment
{
    Center = 0,
    Distributed = 1,
    Left = 2,
    NoControl = 3,
}

public enum XLPhoneticType
{
    FullWidthKatakana = 0,
    HalfWidthKatakana = 1,
    Hiragana = 2,
    NoConversion = 3,
}

public interface IXLPhonetics : IXLFontBase, IEnumerable<IXLPhonetic>, IEquatable<IXLPhonetics>
{
    public IXLPhonetics SetBold();
    public IXLPhonetics SetBold(bool value);
    public IXLPhonetics SetItalic();
    public IXLPhonetics SetItalic(bool value);
    public IXLPhonetics SetUnderline();
    public IXLPhonetics SetUnderline(XLFontUnderlineValues value);
    public IXLPhonetics SetStrikethrough();
    public IXLPhonetics SetStrikethrough(bool value);
    public IXLPhonetics SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
    public IXLPhonetics SetShadow();
    public IXLPhonetics SetShadow(bool value);
    public IXLPhonetics SetFontSize(double value);
    public IXLPhonetics SetFontColor(XLColor value);
    public IXLPhonetics SetFontName(string value);
    public IXLPhonetics SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
    public IXLPhonetics SetFontCharSet(XLFontCharSet value);
    public IXLPhonetics SetFontScheme(XLFontScheme value);

    /// <summary>
    /// Add a phonetic run above a base text. Phonetic runs can't overlap.
    /// </summary>
    /// <param name="text">Text to display above a section of a base text. Can't be empty.</param>
    /// <param name="start">Index of a first character of a base  text above which should <paramref name="text"/> be displayed. Valid values are <c>0</c>..<c>length-1</c>.</param>
    /// <param name="end">The excluded ending index in a base text (the hint is not displayed above the <c>end</c>). Must be &gt; <paramref name="start"/>. Valid values are <c>1</c>..<c>length</c>.</param>
    public IXLPhonetics Add(string text, int start, int end);

    /// <summary>
    /// Remove all phonetic runs. Keeps font properties.
    /// </summary>
    public IXLPhonetics ClearText();

    /// <summary>
    /// Reset font properties to the default font of a container (likely <c>IXLCell</c>). Keeps phonetic runs, <see cref="Type"/> and <see cref="Alignment"/>.
    /// </summary>
    public IXLPhonetics ClearFont();

    /// <summary>
    /// Number of phonetic runs above the base text.
    /// </summary>
    public int Count { get; }

    public XLPhoneticAlignment Alignment { get; set; }
    public XLPhoneticType Type { get; set; }

    public IXLPhonetics SetAlignment(XLPhoneticAlignment phoneticAlignment);
    public IXLPhonetics SetType(XLPhoneticType phoneticType);
}

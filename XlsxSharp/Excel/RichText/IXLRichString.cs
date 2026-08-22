#nullable disable

using System;

namespace XlsxSharp.Excel.RichText;

public interface IXLWithRichString
{
    public IXLRichString AddText(string text);
    public IXLRichString AddNewLine();
}

public interface IXLRichString : IXLFontBase, IEquatable<IXLRichString>, IXLWithRichString
{
    public string Text { get; set; }

    public IXLRichString SetBold();
    public IXLRichString SetBold(bool value);
    public IXLRichString SetItalic();
    public IXLRichString SetItalic(bool value);
    public IXLRichString SetUnderline();
    public IXLRichString SetUnderline(XLFontUnderlineValues value);
    public IXLRichString SetStrikethrough();
    public IXLRichString SetStrikethrough(bool value);
    public IXLRichString SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
    public IXLRichString SetShadow();
    public IXLRichString SetShadow(bool value);
    public IXLRichString SetFontSize(double value);
    public IXLRichString SetFontColor(XLColor value);
    public IXLRichString SetFontName(string value);
    public IXLRichString SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
    public IXLRichString SetFontCharSet(XLFontCharSet value);

    /// <inheritdoc cref="IXLFontBase.FontScheme"/>
    public IXLRichString SetFontScheme(XLFontScheme value);
}

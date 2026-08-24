#nullable disable

namespace XlsxSharp.Excel;

public interface IXLFontBase
{
    public bool Bold { get; set; }

    public bool Italic { get; set; }

    public XLFontUnderlineValues Underline { get; set; }

    public bool Strikethrough { get; set; }

    public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }

    public bool Shadow { get; set; }

    /// <summary>
    /// Get or set font size in points.
    /// </summary>
    /// <value>
    /// Font size is defined in twips and the passed value is rounded to nearest twips. Allowed
    /// values are from 1pt to 409.55pt.
    /// </value>
    /// <exception cref="ArgumentOutOfRangeException">Value is outside of allowed values 1pt to 409.55pt.</exception>
    public double FontSize { get; set; }

    public XLColor FontColor { get; set; }

    public string FontName { get; set; }

    public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; }

    /// <summary>
    /// <para>
    /// Defines an expected character set used by the text of this <c>font</c>. It helps Excel to choose
    /// a font face, either because requested one isn't present or is unsuitable. Each font file contains
    /// a list of charsets it is capable of rendering and this property is used to detect whether the charset
    /// of a text matches the rendering capabilities of a font face and is thus suitable.
    /// </para>
    /// <example>
    /// Example:
    /// The <c>FontCharSet</c> is <c>XLFontCharSet.Default</c>, but the selected font name is <em>B Mitra</em>
    /// that contains only arabic alphabet and declares so in its file. Excel will detect this discrepancy and
    /// choose a different font to display the text. The outcome is that text is not displayed with the <em>B Mitra</em>
    /// font, but with a different one and user doesn't see persian numbers. To use the <em>B Mitra</em> font,
    /// this property must be set to <c>XLFontCharSet.Arabic</c> that would match the font declared capabilities.
    /// </example>
    /// </summary>
    /// <remarks>Due to prevalence of unicode fonts, this property is rarely used.</remarks>
    public XLFontCharSet FontCharSet { get; set; }

    /// <summary>
    /// Determines a theme font scheme a text belongs to. If the text belongs to a scheme and user changes theme
    /// in Excel, the font of the text will switch to the new theme font. Scheme font has precedence and will be
    /// used instead of a set font.
    /// </summary>
    public XLFontScheme FontScheme { get; set; }
}

using System.Xml.Linq;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Reading the pieces of a style that appear outside <c>xl/styles.xml</c> as well as in it.
/// </summary>
internal static class StyleXml
{
    /// <summary>
    /// A <c>CT_Font</c>, which describes a font the same way whether it is a font in the style
    /// table, the properties of a rich text run or the properties of a comment run.
    /// </summary>
    internal static void LoadFont(XElement? fontSource, IXLFontBase fontBase)
    {
        if (fontSource is null)
        {
            return;
        }

        fontBase.Bold = BoolProperty(fontSource, "b");
        fontBase.Italic = BoolProperty(fontSource, "i");
        fontBase.Shadow = BoolProperty(fontSource, "shadow");
        fontBase.Strikethrough = BoolProperty(fontSource, "strike");

        if (fontSource.Element(SpreadsheetXml.Main + "color") is { } fontColor)
        {
            fontBase.FontColor = SpreadsheetXml.ReadColor(fontColor);
        }

        if (
            SpreadsheetXml.Int(fontSource.Element(SpreadsheetXml.Main + "family"), "val") is
            { } family
        )
        {
            fontBase.FontFamilyNumbering = (XLFontFamilyNumberingValues)family;
        }

        if (
            SpreadsheetXml.String(fontSource.Element(SpreadsheetXml.Main + "rFont"), "val") is
            { } fontName
        )
        {
            fontBase.FontName = fontName;
        }

        if (
            SpreadsheetXml.Double(fontSource.Element(SpreadsheetXml.Main + "sz"), "val") is
            { } fontSize
        )
        {
            fontBase.FontSize = fontSize;
        }

        // The three enumerated properties all default to their first value when the element is
        // there but says nothing.
        if (fontSource.Element(SpreadsheetXml.Main + "u") is { } underline)
        {
            fontBase.Underline = SpreadsheetXml.String(underline, "val") is { } underlineValue
                ? StyleXmlEnums.ParseUnderline(underlineValue)
                : XLFontUnderlineValues.Single;
        }

        if (fontSource.Element(SpreadsheetXml.Main + "vertAlign") is { } verticalAlignment)
        {
            fontBase.VerticalAlignment = SpreadsheetXml.String(verticalAlignment, "val")
                is { } verticalAlignmentValue
                ? StyleXmlEnums.ParseVerticalTextAlignment(verticalAlignmentValue)
                : XLFontVerticalTextAlignmentValues.Baseline;
        }

        if (fontSource.Element(SpreadsheetXml.Main + "scheme") is { } scheme)
        {
            fontBase.FontScheme = SpreadsheetXml.String(scheme, "val") is { } schemeValue
                ? StyleXmlEnums.ParseFontScheme(schemeValue)
                : XLFontScheme.None;
        }
    }

    /// <summary>
    /// The <c>CT_BooleanProperty</c> the font flags use: absent is off, present is on, and an
    /// explicit val decides.
    /// </summary>
    private static bool BoolProperty(XElement fontSource, string name) =>
        fontSource.Element(SpreadsheetXml.Main + name) is { } property
        && (SpreadsheetXml.Bool(property, "val") ?? true);
}

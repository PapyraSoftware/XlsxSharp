using System.Globalization;
using System.Xml.Linq;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// The namespaces of a spreadsheet package and the attribute readers every part shares.
/// </summary>
internal static class SpreadsheetXml
{
    internal static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// The Excel 2010 extensions, which carry sparklines and the newer conditional formats.
    /// </summary>
    internal static readonly XNamespace X14 =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    /// <summary>
    /// The shared Excel namespace the 2010 extensions reach into for cell references and
    /// formulas, usually written with the <c>xm</c> prefix.
    /// </summary>
    internal static readonly XNamespace Xm = "http://schemas.microsoft.com/office/excel/2006/main";

    /// <summary>
    /// OOXML booleans are written as 1/0 or true/false, and both have to be accepted.
    /// </summary>
    /// <remarks>
    /// A value that cannot be read reads as absent rather than as an error, which is what the
    /// SDK's typed values did: an unparseable attribute left HasValue false and the caller took
    /// its default. Files in the wild rely on it - one of the test workbooks carries
    /// activeTab="-1", and the sheet it does not point at is meant to fall back to the first.
    /// </remarks>
    internal static bool? Bool(XElement? element, string name) =>
        element?.Attribute(name)?.Value switch
        {
            "1" or "true" or "on" or "True" => true,
            "0" or "false" or "off" or "False" => false,
            _ => null,
        };

    /// <inheritdoc cref="Bool"/>
    internal static uint? UInt(XElement? element, string name) =>
        uint.TryParse(
            element?.Attribute(name)?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out uint parsed
        )
            ? parsed
            : null;

    /// <inheritdoc cref="Bool"/>
    internal static int? Int(XElement? element, string name) =>
        int.TryParse(
            element?.Attribute(name)?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out int parsed
        )
            ? parsed
            : null;

    /// <inheritdoc cref="Bool"/>
    internal static double? Double(XElement? element, string name) =>
        double.TryParse(
            element?.Attribute(name)?.Value,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double parsed
        )
            ? parsed
            : null;

    internal static string? String(XElement? element, string name) =>
        element?.Attribute(name)?.Value;

    /// <summary>
    /// The text of a child element, which OOXML uses where an attribute would have done - the
    /// page breaks and the header and footer strings are written that way.
    /// </summary>
    /// <summary>
    /// The <c>CT_Color</c> shape shared by the main and the x14 namespaces.
    /// </summary>
    /// <remarks>
    /// A color that names none of rgb, indexed or theme - or an index outside the palette - is
    /// automatic, which is how the SDK conversion this replaces read it.
    /// </remarks>
    internal static XLColor ReadColor(XElement? element)
    {
        if (element is null)
        {
            return XLColor.Automatic;
        }

        if (String(element, "rgb") is { } rgb)
        {
            return XLColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
        }

        if (UInt(element, "indexed") is { } indexed && indexed <= 64)
        {
            return XLColor.FromIndex((int)indexed);
        }

        if (UInt(element, "theme") is { } theme)
        {
            return Double(element, "tint") is { } tint
                ? XLColor.FromTheme((XLThemeColor)theme, tint)
                : XLColor.FromTheme((XLThemeColor)theme);
        }

        return XLColor.Automatic;
    }

    internal static string? ElementText(XElement? element, string name) =>
        element?.Element(Main + name)?.Value;
}

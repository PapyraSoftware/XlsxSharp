using System.Globalization;
using System.Xml.Linq;
using XlsxSharp.Extensions;
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
    /// A new root element in the main namespace, declared as the default namespace - the way
    /// Excel writes its own parts.
    /// </summary>
    internal static XElement NewRoot(string localName, params object?[] content) =>
        new(Main + localName, new XAttribute("xmlns", Main.NamespaceName), content);

    /// <summary>
    /// Declares <paramref name="ns"/> on <paramref name="root"/> under <paramref name="prefix"/>,
    /// unless the root already declares it - under whatever prefix the part was loaded with.
    /// </summary>
    internal static void EnsureDeclared(XElement root, string prefix, XNamespace ns)
    {
        if (!root.Attributes().Any(a => a.IsNamespaceDeclaration && a.Value == ns.NamespaceName))
        {
            root.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
        }
    }

    /// <summary>
    /// An <c>xsd:boolean</c> attribute: <c>true</c>, <c>false</c>, <c>1</c> or <c>0</c>, and
    /// nothing else.
    /// </summary>
    /// <remarks>
    /// A value that is not one of those - including the <c>on</c>/<c>off</c> and <c>True</c> forms
    /// the schema does not allow - reads as absent rather than as an error, and the caller takes
    /// its default. Files in the wild rely on that leniency for out-of-range values: one of the
    /// test workbooks carries activeTab="-1", and the sheet it does not point at is meant to fall
    /// back to the first.
    /// </remarks>
    internal static bool? Bool(XElement? element, string name) =>
        ParseBoolean(element?.Attribute(name)?.Value);

    /// <inheritdoc cref="Bool"/>
    internal static bool? ParseBoolean(string? value) =>
        value?.Trim(XmlWhitespace) switch
        {
            "1" or "true" => true,
            "0" or "false" => false,
            _ => null,
        };

    /// <summary>The characters <c>xsd:whiteSpace="collapse"</c> strips around a value.</summary>
    private static readonly char[] XmlWhitespace = [' ', '\t', '\n', '\r'];

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
    /// automatic.
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

    /// <summary>
    /// Writes a colour into a <c>CT_Color</c>, in whichever of the three ways the workbook model
    /// holds it. An automatic colour says nothing at all.
    /// </summary>
    /// <param name="isDifferential">
    /// A differential format leaves index 64 - the transparent one - unsaid.
    /// </param>
    internal static void SetColor(XElement element, XLColor color, bool isDifferential = false)
    {
        switch (color.ColorType)
        {
            case XLColorType.Color:
                element.SetAttributeValue("rgb", color.Color.ToHex());
                break;

            case XLColorType.Indexed:
                if (!isDifferential || color.Indexed != 64)
                {
                    element.SetAttributeValue("indexed", (uint)color.Indexed);
                }

                break;

            case XLColorType.Theme:
                element.SetAttributeValue("theme", (uint)color.ThemeColor);
                if (color.ThemeTint != 0)
                {
                    element.SetAttributeValue(
                        "tint",
                        color.ThemeTint.ToString(CultureInfo.InvariantCulture)
                    );
                }

                break;
        }
    }

    internal static string? ElementText(XElement? element, string name) =>
        element?.Element(Main + name)?.Value;
}

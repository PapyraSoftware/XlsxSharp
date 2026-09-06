using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writing <c>xl/theme/theme1.xml</c>.
/// </summary>
/// <remarks>
/// The theme is the stock Office theme, and every part of it but its twelve colours is the same
/// in every workbook XlsxSharp writes. It is kept as the document it is, in OfficeTheme.xml
/// beside this file, rather than as the several hundred statements it takes to build one - the
/// gradient stops and the effect lists say nothing that reading them in XML does not.
/// </remarks>
internal static class ThemePartWriter
{
    private const string ResourceName = "XlsxSharp.Excel.IO.OfficeTheme.xml";

    private static readonly XNamespace Drawing =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static void GenerateContent(OpcPart themePart, XLTheme theme)
    {
        XDocument document = ReadTemplate();
        XElement colorScheme =
            document.Root?.Element(Drawing + "themeElements")?.Element(Drawing + "clrScheme")
            ?? throw new InvalidOperationException($"{ResourceName} has no colour scheme.");

        // The two system colours name the window and its text, and carry the colour they last
        // resolved to; the rest are plain RGB.
        SetSystemColor("dk1", theme.Text1);
        SetSystemColor("lt1", theme.Background1);
        SetColor("dk2", theme.Text2);
        SetColor("lt2", theme.Background2);
        SetColor("accent1", theme.Accent1);
        SetColor("accent2", theme.Accent2);
        SetColor("accent3", theme.Accent3);
        SetColor("accent4", theme.Accent4);
        SetColor("accent5", theme.Accent5);
        SetColor("accent6", theme.Accent6);
        SetColor("hlink", theme.Hyperlink);
        SetColor("folHlink", theme.FollowedHyperlink);

        using Stream partStream = themePart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { CloseOutput = true, Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );
        document.Save(xml);

        void SetSystemColor(string name, XLColor color) =>
            Color(name, Drawing + "sysClr").SetAttributeValue("lastClr", Rgb(color));

        void SetColor(string name, XLColor color) =>
            Color(name, Drawing + "srgbClr").SetAttributeValue("val", Rgb(color));

        XElement Color(string name, XName colorName) =>
            colorScheme.Element(Drawing + name)?.Element(colorName)
            ?? throw new InvalidOperationException($"{ResourceName} has no {name} colour.");
    }

    /// <summary>
    /// The colour without its alpha, which the theme has no place for.
    /// </summary>
    private static string Rgb(XLColor color) => color.Color.ToHex()[2..];

    private static XDocument ReadTemplate()
    {
        using Stream stream =
            Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException($"{ResourceName} is missing.");

        // The template is indented to be read; none of that indentation belongs in the part.
        using XmlReader reader = XmlReader.Create(
            stream,
            new XmlReaderSettings { IgnoreWhitespace = true }
        );
        return XDocument.Load(reader);
    }
}

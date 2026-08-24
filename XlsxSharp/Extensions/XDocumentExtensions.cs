using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace XlsxSharp.Extensions;

internal static class XDocumentExtensions
{
    /// <summary>
    /// Serialize a node the way <see cref="XNode.ToString()"/> does, but with the line break spelled
    /// out. The default of <c>XmlWriterSettings</c> follows <c>Environment.NewLine</c>, which would
    /// put LF into a part on unix and CRLF on windows, so the same workbook would be saved to files
    /// that differ byte wise depending on where the code runs. CRLF is what Excel writes.
    /// </summary>
    public static string ToXmlString(this XNode node)
    {
        XmlWriterSettings settings = new()
        {
            Indent = true,
            OmitXmlDeclaration = true,
            // A document writes a start of document, which a fragment level writer rejects, while an
            // element on its own is not a whole document. This is the same split XNode.ToString makes.
            ConformanceLevel =
                node is XDocument ? ConformanceLevel.Document : ConformanceLevel.Fragment,
            NewLineChars = "\r\n",
        };

        using StringWriter stringWriter = new(CultureInfo.InvariantCulture);
        using (XmlWriter xmlWriter = XmlWriter.Create(stringWriter, settings))
        {
            node.WriteTo(xmlWriter);
        }

        return stringWriter.ToString();
    }

    public static XDocument? Load(Stream stream)
    {
        using (XmlReader reader = XmlReader.Create(stream))
        {
            try
            {
                return XDocument.Load(reader);
            }
            catch (XmlException)
            {
                return null;
            }
        }
    }
}

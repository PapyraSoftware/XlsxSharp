// Keep this file CodeMaid organised and cleaned

using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace XlsxSharp.Extensions;

internal static class XDocumentExtensions
{
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

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// Compares a single element from an <see cref="XDocument"/> query against the supplied XML.
/// </summary>
internal static class XmlAssert
{
    public static void MatchesXml(IEnumerable<XElement> elements, string xml)
    {
        XElement element = elements.Single();
        XDocument expected = XDocument.Load(new StringReader(xml));
        if (!element.SemanticallyEqual(expected.Root))
        {
            ClassicAssert.Fail($"XML should semantically match {xml}, but was {element}.");
        }
    }
}

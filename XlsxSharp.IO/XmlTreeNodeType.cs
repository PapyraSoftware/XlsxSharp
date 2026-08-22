using System.Xml;

namespace XlsxSharp.IO;

/// <summary>
/// A node the <see cref="IXmlReader"/> is currently on.
/// </summary>
public enum XmlTreeNodeType
{
    /// <summary>
    /// Reader hasn't yet read anything or has already read everything.
    /// </summary>
    None,

    /// <summary>
    /// Reader is on a opening element.
    /// </summary>
    OpenElement,

    /// <summary>
    /// Reader is on a closing element.
    /// </summary>
    CloseElement,

    /// <summary>
    /// Reader is on a text node. There can be multiple subsequent text nodes. For example this:
    /// <c>First text node &lt;![CDATA[Second text node]]&gt;</c>
    /// will produce two subsequent text nodes: "First text node" and "Second text node".
    /// </summary>
    /// <remarks>
    /// If the underlaying <see cref="XmlReader"/> has <see cref="XmlReaderSettings.IgnoreWhitespace"/>
    /// set to <c>false</c>, even insignificant whitespaces will produce text nodes.
    /// </remarks>
    Text,
}

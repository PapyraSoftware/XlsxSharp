using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// The content of a shared string (<c>si</c>) or an inline string (<c>is</c>), a <c>CT_Rst</c>,
/// read forward-only off the stream.
/// </summary>
/// <remarks>
/// Only what a cell needs is kept: the text, the runs with their formatting and the phonetic
/// guide. A run's <c>rPr</c> and the <c>phoneticPr</c> stay elements, because the font reader
/// takes one; they are small and only exist for formatted text.
/// </remarks>
internal sealed class StringItem
{
    private static readonly string Main = OoxmlConst.Main2006SsNs;

    private StringItem() { }

    /// <summary>The content of the <c>t</c> element, or <c>null</c> when there is none.</summary>
    internal string? Text { get; private set; }

    /// <summary>The formatted runs, or <c>null</c> for text without any.</summary>
    internal List<(XElement? Properties, string Text)>? Runs { get; private set; }

    /// <summary>The <c>phoneticPr</c> element, or <c>null</c>.</summary>
    internal XElement? PhoneticProperties { get; private set; }

    /// <summary>The <c>rPh</c> runs of the phonetic guide, or <c>null</c>.</summary>
    internal List<(string Text, int Start, int End)>? Phonetics { get; private set; }

    /// <summary>Whether this is just text: no runs and no phonetic guide.</summary>
    internal bool IsPlain =>
        this.Runs is null && this.PhoneticProperties is null && this.Phonetics is null;

    /// <summary>
    /// Reads the element the reader is on and leaves the reader on whatever follows it.
    /// </summary>
    internal static StringItem Read(XmlReader reader)
    {
        StringItem item = new();
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return item;
        }

        reader.ReadStartElement();
        while (reader.NodeType != XmlNodeType.EndElement)
        {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != Main)
            {
                reader.Skip();
                continue;
            }

            switch (reader.LocalName)
            {
                case "t":
                    item.Text = reader.ReadElementContentAsString();
                    break;
                case "r":
                    (item.Runs ??= []).Add(ReadRun(reader));
                    break;
                case "rPh":
                    (item.Phonetics ??= []).Add(ReadPhonetic(reader));
                    break;
                case "phoneticPr":
                    item.PhoneticProperties = (XElement)XNode.ReadFrom(reader);
                    break;
                default:
                    reader.Skip();
                    break;
            }
        }

        reader.ReadEndElement();
        return item;
    }

    private static (XElement? Properties, string Text) ReadRun(XmlReader reader)
    {
        XElement? properties = null;
        string text = string.Empty;
        ReadChildren(
            reader,
            () =>
            {
                switch (reader.LocalName)
                {
                    case "rPr":
                        properties = (XElement)XNode.ReadFrom(reader);
                        return;
                    case "t":
                        text = reader.ReadElementContentAsString();
                        return;
                    default:
                        reader.Skip();
                        return;
                }
            }
        );

        return (properties, text);
    }

    private static (string Text, int Start, int End) ReadPhonetic(XmlReader reader)
    {
        int start = RequiredInt(reader, "sb");
        int end = RequiredInt(reader, "eb");
        string text = string.Empty;
        ReadChildren(
            reader,
            () =>
            {
                if (reader.LocalName == "t")
                {
                    text = reader.ReadElementContentAsString();
                }
                else
                {
                    reader.Skip();
                }
            }
        );

        return (text, start, end);
    }

    /// <summary>
    /// Calls <paramref name="readChild"/> for every child element in the main namespace, which
    /// has to consume it, and skips everything else.
    /// </summary>
    private static void ReadChildren(XmlReader reader, Action readChild)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return;
        }

        reader.ReadStartElement();
        while (reader.NodeType != XmlNodeType.EndElement)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == Main)
            {
                readChild();
            }
            else
            {
                reader.Skip();
            }
        }

        reader.ReadEndElement();
    }

    private static int RequiredInt(XmlReader reader, string name) =>
        uint.TryParse(
            reader.GetAttribute(name),
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out uint value
        )
            ? checked((int)value)
            : throw PartStructureException.MissingAttribute(name);
}

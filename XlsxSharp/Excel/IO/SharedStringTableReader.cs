using System.Globalization;
using System.Xml;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// One entry of the shared string table: the decoded text of a plain entry, or the whole item
/// of one with runs or a phonetic guide.
/// </summary>
internal readonly record struct SharedString(string Text, StringItem? Item);

/// <summary>
/// Reads <c>xl/sharedStrings.xml</c> forward-only. A plain entry - nearly every entry of a real
/// workbook - is kept as its decoded string and nothing else, rather than as the element it was
/// read from.
/// </summary>
internal static class SharedStringTableReader
{
    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        CloseInput = false,
    };

    internal static SharedString[] Read(Stream stream)
    {
        using XmlReader reader = XmlReader.Create(stream, ReaderSettings);
        reader.MoveToContent();
        if (
            reader.NodeType != XmlNodeType.Element
            || reader.LocalName != "sst"
            || reader.NamespaceURI != OoxmlConst.Main2006SsNs
        )
        {
            throw PartStructureException.ExpectedElementNotFound("sst");
        }

        List<SharedString> entries = new(Capacity(reader.GetAttribute("uniqueCount")));
        if (reader.IsEmptyElement)
        {
            return [];
        }

        reader.ReadStartElement();
        while (reader.NodeType != XmlNodeType.EndElement && !reader.EOF)
        {
            if (
                reader.NodeType == XmlNodeType.Element
                && reader.LocalName == "si"
                && reader.NamespaceURI == OoxmlConst.Main2006SsNs
            )
            {
                StringItem item = StringItem.Read(reader);
                entries.Add(
                    item.IsPlain
                        ? new SharedString(XStringConvert.Decode(item.Text) ?? string.Empty, null)
                        : new SharedString(string.Empty, item)
                );
            }
            else
            {
                reader.Skip();
            }
        }

        return [.. entries];
    }

    /// <summary>
    /// The list is sized from <c>uniqueCount</c> when the part states it, within reason - the
    /// attribute is only a hint and a corrupt one must not allocate gigabytes up front.
    /// </summary>
    private static int Capacity(string? uniqueCount) =>
        int.TryParse(uniqueCount, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)
            ? Math.Clamp(count, 0, 1 << 20)
            : 0;
}

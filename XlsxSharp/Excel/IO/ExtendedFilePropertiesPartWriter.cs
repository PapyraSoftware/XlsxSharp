#nullable disable

using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writes <c>/docProps/app.xml</c>.
/// </summary>
/// <remarks>
/// The part is patched rather than rewritten. XlsxSharp owns only a handful of the elements the
/// extended properties schema allows, and a workbook that came from Excel carries others -
/// AppVersion, LinksUpToDate, HyperlinksChanged - that have to survive a load and save. So the
/// existing document is read, the owned elements are replaced in place, and everything else is
/// left where and as it was.
/// </remarks>
internal class ExtendedFilePropertiesPartWriter
{
    private static readonly XNamespace Ap =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static readonly XNamespace Vt =
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>
    /// The order the owned elements are created in when the part does not have them yet. It is
    /// the order the SDK produced, which the reference workbooks are recorded with.
    /// </summary>
    private static readonly string[] OwnedOrder =
    [
        "Application",
        "Company",
        "Manager",
        "TitlesOfParts",
        "HeadingPairs",
        "DocSecurity",
        "ScaleCrop",
    ];

    internal static void GenerateContent(OpcPart extendedFilePropertiesPart, XLWorkbook workbook)
    {
        XDocument document = ReadExisting(extendedFilePropertiesPart);
        XElement properties = document.Root;

        List<XLWorksheet> worksheets =
        [
            .. ((IEnumerable<XLWorksheet>)workbook.WorksheetsInternal).OrderBy(w => w.Position),
        ];

        List<string> namedRanges = GetModifiedNamedRanges(workbook);

        // These three are filled in only when the part does not have them. A workbook that came
        // from another producer says so - "Microsoft Access" for one of the test files - and
        // overwriting that would be a lie about where the file came from.
        SetIfMissing(properties, "Application", "Microsoft Excel");
        SetIfMissing(properties, "DocSecurity", "0");
        SetIfMissing(properties, "ScaleCrop", "false");

        // The two vectors, on the other hand, describe the workbook as it is being saved now and
        // are always rewritten.
        SetOwned(
            properties,
            "TitlesOfParts",
            new XElement(
                Ap + "TitlesOfParts",
                new XElement(
                    Vt + "vector",
                    new XAttribute("baseType", "lpstr"),
                    new XAttribute("size", worksheets.Count + namedRanges.Count),
                    worksheets.Select(w => new XElement(Vt + "lpstr", w.Name)),
                    namedRanges.Select(nr => new XElement(Vt + "lpstr", nr))
                )
            )
        );

        SetOwned(
            properties,
            "HeadingPairs",
            new XElement(
                Ap + "HeadingPairs",
                new XElement(
                    Vt + "vector",
                    new XAttribute("baseType", "variant"),
                    new XAttribute("size", 4),
                    HeadingPair("Worksheets", worksheets.Count),
                    HeadingPair("Named Ranges", namedRanges.Count)
                )
            )
        );

        // Manager and Company are the only two the workbook model can clear: a null property
        // means "leave whatever the part had", an empty one means "remove it".
        SetOptional(properties, "Manager", workbook.Properties.Manager);
        SetOptional(properties, "Company", workbook.Properties.Company);

        using Stream partStream = extendedFilePropertiesPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );

        document.Save(xml);
    }

    /// <summary>
    /// The existing document normalised, or a new one when the part is empty.
    /// </summary>
    /// <remarks>
    /// Normalising matters for byte fidelity. A workbook written by Excel puts the extended
    /// properties in a default namespace with unprefixed elements; what XlsxSharp has always
    /// written back is the same content re-serialised with the "ap" prefix on every element and
    /// the two prefixes declared on the root, vt first. So the loaded elements are moved into a
    /// root of that shape and their own namespace declarations dropped, which leaves the prefix
    /// to be resolved from the new root.
    /// </remarks>
    private static XDocument ReadExisting(OpcPart part)
    {
        XElement loaded = null;
        bool standalone = false;

        if (part.Length > 0)
        {
            using Stream stream = part.GetReadStream();
            try
            {
                XDocument existing = XDocument.Load(stream);
                loaded = existing.Root;
                standalone = string.Equals(
                    existing.Declaration?.Standalone,
                    "yes",
                    StringComparison.OrdinalIgnoreCase
                );
            }
            catch (XmlException)
            {
                // An unreadable app.xml is not worth failing a save over, it holds no data
                // the workbook depends on. Start over instead.
            }
        }

        XElement properties = new(
            Ap + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", Vt.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "ap", Ap.NamespaceName)
        );

        if (loaded is not null)
        {
            foreach (XElement child in loaded.Elements())
            {
                XElement copy = new(child);
                copy.DescendantsAndSelf()
                    .Attributes()
                    .Where(a => a.IsNamespaceDeclaration)
                    .ToList()
                    .ForEach(a => a.Remove());

                properties.Add(copy);
            }
        }

        return standalone
            ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), properties)
            : new XDocument(properties);
    }

    /// <summary>
    /// A heading pair is two sibling variants, the name and then the count - not one variant
    /// holding both.
    /// </summary>
    private static XElement[] HeadingPair(string name, int count) =>
        [
            new XElement(Vt + "variant", new XElement(Vt + "lpstr", name)),
            new XElement(Vt + "variant", new XElement(Vt + "i4", count.ToInvariantString())),
        ];

    /// <summary>
    /// Replaces an element XlsxSharp owns, keeping its position when the part already had it and
    /// otherwise putting it where <see cref="OwnedOrder"/> says.
    /// </summary>
    private static void SetOwned(XElement properties, string localName, XElement replacement)
    {
        XElement existing = properties.Element(Ap + localName);
        if (existing is not null)
        {
            existing.ReplaceWith(replacement);
            return;
        }

        InsertInOwnedOrder(properties, localName, replacement);
    }

    /// <summary>
    /// Adds an element only when the part does not already have one, leaving whatever a previous
    /// producer wrote alone.
    /// </summary>
    private static void SetIfMissing(XElement properties, string localName, string value)
    {
        if (properties.Element(Ap + localName) is null)
        {
            InsertInOwnedOrder(properties, localName, new XElement(Ap + localName, value));
        }
    }

    private static void SetOptional(XElement properties, string localName, string value)
    {
        if (value is null)
        {
            return;
        }

        XElement existing = properties.Element(Ap + localName);
        if (string.IsNullOrWhiteSpace(value))
        {
            existing?.Remove();
            return;
        }

        if (existing is not null)
        {
            existing.Value = value;
            return;
        }

        InsertInOwnedOrder(properties, localName, new XElement(Ap + localName, value));
    }

    /// <summary>
    /// Puts a new element after the last owned element that precedes it in
    /// <see cref="OwnedOrder"/>, so that a part built from nothing comes out in that order and a
    /// part that already has content keeps the elements it has where they are.
    /// </summary>
    private static void InsertInOwnedOrder(XElement properties, string localName, XElement element)
    {
        int position = Array.IndexOf(OwnedOrder, localName);

        XElement predecessor = null;
        for (int i = 0; i < position; i++)
        {
            XElement candidate = properties.Element(Ap + OwnedOrder[i]);
            if (candidate is not null)
            {
                predecessor = candidate;
            }
        }

        if (predecessor is null)
        {
            properties.AddFirst(element);
        }
        else
        {
            predecessor.AddAfterSelf(element);
        }
    }

    private static List<string> GetModifiedNamedRanges(XLWorkbook workbook)
    {
        List<string> namedRanges = [];
        foreach (XLWorksheet sheet in workbook.WorksheetsInternal)
        {
            namedRanges.AddRange(
                sheet.DefinedNames.Select<XLDefinedName, string>(n => sheet.Name + "!" + n.Name)
            );
            namedRanges.Add(sheet.Name + "!Print_Area");
            namedRanges.Add(sheet.Name + "!Print_Titles");
        }

        namedRanges.AddRange(
            workbook.DefinedNamesInternal.Select<XLDefinedName, string>(n => n.Name)
        );
        return namedRanges;
    }
}

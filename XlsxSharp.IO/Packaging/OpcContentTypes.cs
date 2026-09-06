using System.Xml;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// The <c>[Content_Types].xml</c> stream of a package (ECMA-376 Part 2 §10.1.2). It maps every
/// part to a content type, either through a default for the part's extension or through an
/// override for the individual part.
/// </summary>
internal sealed class OpcContentTypes
{
    internal const string PartName = "/[Content_Types].xml";

    private const string Ns = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>Extension (lower case, without the dot) to content type.</summary>
    private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Part name to content type, taking precedence over <see cref="_defaults"/>.</summary>
    private readonly Dictionary<string, string> _overrides = new(OpcPartName.Comparer);

    internal static OpcContentTypes CreateEmpty()
    {
        OpcContentTypes contentTypes = new();

        // Every package needs the relationship default; xml is what Excel writes for the parts
        // that have no more specific type.
        contentTypes._defaults["rels"] = OpcContentType.Relationships;
        contentTypes._defaults["xml"] = OpcContentType.Xml;
        return contentTypes;
    }

    internal static OpcContentTypes Read(Stream stream)
    {
        OpcContentTypes contentTypes = new();

        XmlReaderSettings settings = new()
        {
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            DtdProcessing = DtdProcessing.Prohibit,
        };

        using XmlReader reader = XmlReader.Create(stream, settings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != Ns)
            {
                continue;
            }

            switch (reader.LocalName)
            {
                case "Default":
                {
                    string? extension = reader.GetAttribute("Extension");
                    string? contentType = reader.GetAttribute("ContentType");
                    if (extension is not null && contentType is not null)
                    {
                        contentTypes._defaults[extension.TrimStart('.')] = contentType;
                    }

                    break;
                }

                case "Override":
                {
                    string? partName = reader.GetAttribute("PartName");
                    string? contentType = reader.GetAttribute("ContentType");
                    if (partName is not null && contentType is not null)
                    {
                        contentTypes._overrides[OpcPartName.Normalize(partName)] = contentType;
                    }

                    break;
                }
            }
        }

        return contentTypes;
    }

    /// <summary>
    /// The content type of a part, or <c>null</c> when the package declares none for it.
    /// </summary>
    internal string? GetContentType(string partName)
    {
        if (this._overrides.TryGetValue(partName, out string? contentType))
        {
            return contentType;
        }

        string extension = OpcPartName.GetExtension(partName);
        return
            extension.Length > 0 && this._defaults.TryGetValue(extension, out string? byExtension)
            ? byExtension
            : null;
    }

    /// <summary>
    /// Declares the content type of a part. An existing default for the part's extension is
    /// reused when it already says the right thing, otherwise an override is written, which is
    /// always valid.
    /// </summary>
    internal void SetContentType(string partName, string contentType)
    {
        string extension = OpcPartName.GetExtension(partName);

        if (
            extension.Length > 0
            && this._defaults.TryGetValue(extension, out string? byExtension)
            && string.Equals(byExtension, contentType, StringComparison.OrdinalIgnoreCase)
        )
        {
            this._overrides.Remove(partName);
            return;
        }

        // A package with no default for an extension needs one for the binary parts, because
        // consumers reject a package where a part has no content type at all. For everything
        // else an override is the safer choice: xml parts share an extension but not a type.
        if (extension.Length > 0 && !this._defaults.ContainsKey(extension) && !IsXml(contentType))
        {
            this._defaults[extension] = contentType;
            return;
        }

        this._overrides[partName] = contentType;
    }

    internal void Remove(string partName) => this._overrides.Remove(partName);

    internal void Write(Stream stream)
    {
        XmlWriterSettings settings = new() { CloseOutput = false, Encoding = OpcXml.NoBomUtf8 };

        using XmlWriter writer = XmlWriter.Create(stream, settings);
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("Types", Ns);

        foreach ((string extension, string contentType) in this._defaults.OrderBy(x => x.Key))
        {
            writer.WriteStartElement("Default", Ns);
            writer.WriteAttributeString("Extension", extension);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        foreach ((string partName, string contentType) in this._overrides.OrderBy(x => x.Key))
        {
            writer.WriteStartElement("Override", Ns);
            writer.WriteAttributeString("PartName", partName);
            writer.WriteAttributeString("ContentType", contentType);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static bool IsXml(string contentType) =>
        contentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase);
}

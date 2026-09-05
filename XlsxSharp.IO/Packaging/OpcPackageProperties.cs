using System.Globalization;
using System.Xml;

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// The core properties of a package, i.e. the content of the part related from the package with
/// <see cref="OpcRelationshipType.CoreProperties"/>, conventionally <c>/docProps/core.xml</c>
/// (ECMA-376 Part 2 §11).
/// </summary>
/// <remarks>
/// A property that is not present in the part reads as <c>null</c>, which is how a consumer tells
/// "not set" from "set to the empty string". Setting one marks the properties dirty, and the part
/// is rewritten when the package is saved.
/// </remarks>
public sealed class OpcPackageProperties
{
    private const string CoreNs =
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    private const string DcNs = "http://purl.org/dc/elements/1.1/";
    private const string DcTermsNs = "http://purl.org/dc/terms/";
    private const string DcmiTypeNs = "http://purl.org/dc/dcmitype/";
    private const string XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

    /// <summary>Where the part is created when the package does not have one yet.</summary>
    internal const string DefaultPartName = "/docProps/core.xml";

    private string? _category;
    private string? _contentStatus;
    private string? _contentType;
    private DateTime? _created;
    private string? _creator;
    private string? _description;
    private string? _identifier;
    private string? _keywords;
    private string? _language;
    private string? _lastModifiedBy;
    private DateTime? _lastPrinted;
    private DateTime? _modified;
    private string? _revision;
    private string? _subject;
    private string? _title;
    private string? _version;

    internal OpcPackageProperties() { }

    /// <summary>Whether a setter ran, so that the part has to be rewritten on save.</summary>
    internal bool IsDirty { get; private set; }

    public string? Category
    {
        get => this._category;
        set => this.Set(ref this._category, value);
    }

    public string? ContentStatus
    {
        get => this._contentStatus;
        set => this.Set(ref this._contentStatus, value);
    }

    public string? ContentType
    {
        get => this._contentType;
        set => this.Set(ref this._contentType, value);
    }

    public DateTime? Created
    {
        get => this._created;
        set => this.Set(ref this._created, value);
    }

    public string? Creator
    {
        get => this._creator;
        set => this.Set(ref this._creator, value);
    }

    public string? Description
    {
        get => this._description;
        set => this.Set(ref this._description, value);
    }

    public string? Identifier
    {
        get => this._identifier;
        set => this.Set(ref this._identifier, value);
    }

    public string? Keywords
    {
        get => this._keywords;
        set => this.Set(ref this._keywords, value);
    }

    public string? Language
    {
        get => this._language;
        set => this.Set(ref this._language, value);
    }

    public string? LastModifiedBy
    {
        get => this._lastModifiedBy;
        set => this.Set(ref this._lastModifiedBy, value);
    }

    public DateTime? LastPrinted
    {
        get => this._lastPrinted;
        set => this.Set(ref this._lastPrinted, value);
    }

    public DateTime? Modified
    {
        get => this._modified;
        set => this.Set(ref this._modified, value);
    }

    public string? Revision
    {
        get => this._revision;
        set => this.Set(ref this._revision, value);
    }

    public string? Subject
    {
        get => this._subject;
        set => this.Set(ref this._subject, value);
    }

    public string? Title
    {
        get => this._title;
        set => this.Set(ref this._title, value);
    }

    public string? Version
    {
        get => this._version;
        set => this.Set(ref this._version, value);
    }

    internal static OpcPackageProperties Read(Stream stream)
    {
        OpcPackageProperties properties = new();

        using XmlReader reader = XmlReader.Create(stream, OpcXml.ReaderSettings);
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            string ns = reader.NamespaceURI;
            string name = reader.LocalName;
            if (ns == CoreNs && name == "coreProperties")
            {
                reader.Read();
                continue;
            }

            // ReadElementContentAsString consumes the element and leaves the reader on the node
            // after it, so this loop must not advance again on this iteration.
            string value = reader.ReadElementContentAsString();

            switch (ns)
            {
                case DcNs:
                    switch (name)
                    {
                        case "creator":
                            properties._creator = value;
                            break;
                        case "description":
                            properties._description = value;
                            break;
                        case "identifier":
                            properties._identifier = value;
                            break;
                        case "language":
                            properties._language = value;
                            break;
                        case "subject":
                            properties._subject = value;
                            break;
                        case "title":
                            properties._title = value;
                            break;
                    }

                    break;

                case DcTermsNs:
                    switch (name)
                    {
                        case "created":
                            properties._created = ParseDate(value);
                            break;
                        case "modified":
                            properties._modified = ParseDate(value);
                            break;
                    }

                    break;

                case CoreNs:
                    switch (name)
                    {
                        case "category":
                            properties._category = value;
                            break;
                        case "contentStatus":
                            properties._contentStatus = value;
                            break;
                        case "contentType":
                            properties._contentType = value;
                            break;
                        case "keywords":
                            properties._keywords = value;
                            break;
                        case "lastModifiedBy":
                            properties._lastModifiedBy = value;
                            break;
                        case "lastPrinted":
                            properties._lastPrinted = ParseDate(value);
                            break;
                        case "revision":
                            properties._revision = value;
                            break;
                        case "version":
                            properties._version = value;
                            break;
                    }

                    break;
            }
        }

        return properties;
    }

    internal void Write(Stream stream)
    {
        XmlWriterSettings settings = new() { CloseOutput = false, Encoding = OpcXml.NoBomUtf8 };

        using XmlWriter writer = XmlWriter.Create(stream, settings);
        writer.WriteStartDocument(standalone: true);
        writer.WriteStartElement("cp", "coreProperties", CoreNs);
        writer.WriteAttributeString("xmlns", "dc", null, DcNs);
        writer.WriteAttributeString("xmlns", "dcterms", null, DcTermsNs);
        writer.WriteAttributeString("xmlns", "dcmitype", null, DcmiTypeNs);
        writer.WriteAttributeString("xmlns", "xsi", null, XsiNs);

        // The order follows the sequence in the core properties schema, which is what consumers
        // that validate against it expect.
        WriteText(writer, "dc", "creator", DcNs, this._creator);
        WriteText(writer, "cp", "keywords", CoreNs, this._keywords);
        WriteText(writer, "dc", "description", DcNs, this._description);
        WriteText(writer, "dc", "title", DcNs, this._title);
        WriteText(writer, "dc", "subject", DcNs, this._subject);
        WriteText(writer, "cp", "lastModifiedBy", CoreNs, this._lastModifiedBy);
        WriteDate(writer, "dcterms", "created", DcTermsNs, this._created);
        WriteDate(writer, "dcterms", "modified", DcTermsNs, this._modified);
        WriteDate(writer, "cp", "lastPrinted", CoreNs, this._lastPrinted);
        WriteText(writer, "cp", "contentType", CoreNs, this._contentType);
        WriteText(writer, "cp", "contentStatus", CoreNs, this._contentStatus);
        WriteText(writer, "cp", "category", CoreNs, this._category);
        WriteText(writer, "cp", "version", CoreNs, this._version);
        WriteText(writer, "cp", "revision", CoreNs, this._revision);
        WriteText(writer, "dc", "identifier", DcNs, this._identifier);
        WriteText(writer, "dc", "language", DcNs, this._language);

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteText(
        XmlWriter writer,
        string prefix,
        string localName,
        string ns,
        string? value
    )
    {
        if (value is not null)
        {
            writer.WriteElementString(prefix, localName, ns, value);
        }
    }

    private static void WriteDate(
        XmlWriter writer,
        string prefix,
        string localName,
        string ns,
        DateTime? value
    )
    {
        if (value is null)
        {
            return;
        }

        writer.WriteStartElement(prefix, localName, ns);

        // The W3CDTF type annotation is required on the date properties, §11.2.
        writer.WriteAttributeString("type", XsiNs, "dcterms:W3CDTF");
        writer.WriteString(FormatDate(value.Value));
        writer.WriteEndElement();
    }

    /// <summary>
    /// W3CDTF, which for these properties means an ISO 8601 instant in UTC, as Excel writes them.
    /// </summary>
    private static string FormatDate(DateTime value) =>
        (
            value.Kind == DateTimeKind.Unspecified
                ? DateTime.SpecifyKind(value, DateTimeKind.Utc)
                : value.ToUniversalTime()
        ).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

    private static DateTime? ParseDate(string value) =>
        DateTime.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
            out DateTime parsed
        )
            ? parsed
            : null;

    private void Set<T>(ref T field, T value)
    {
        field = value;
        this.IsDirty = true;
    }
}

namespace XlsxSharp.IO.Packaging;

/// <summary>
/// Content types of the package level parts. SpreadsheetML part types live with the code that
/// writes those parts, these are the ones the packaging layer itself has to know about.
/// </summary>
public static class OpcContentType
{
    /// <summary>Content type of a <c>.rels</c> part.</summary>
    public const string Relationships = "application/vnd.openxmlformats-package.relationships+xml";

    /// <summary>Content type Excel uses as the default for the <c>xml</c> extension.</summary>
    public const string Xml = "application/xml";

    /// <summary>Content type of <c>/docProps/core.xml</c>.</summary>
    public const string CoreProperties =
        "application/vnd.openxmlformats-package.core-properties+xml";
}

/// <summary>
/// Relationship types of the package level parts.
/// </summary>
public static class OpcRelationshipType
{
    /// <summary>Relationship type of <c>/docProps/core.xml</c>.</summary>
    public const string CoreProperties =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
}

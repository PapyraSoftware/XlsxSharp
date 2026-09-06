using System.Reflection;
using System.Xml;
using System.Xml.Schema;

namespace XlsxSharp.Excel.IO.Schemas;

/// <summary>
/// The OOXML schemas <see cref="SchemaValidator"/> validates saved parts against. See
/// <c>PROVENANCE.md</c> next to the embedded <c>.xsd</c> files for what they are and where they
/// come from.
/// </summary>
internal static class OoxmlSchemas
{
    /// <summary>
    /// Every embedded schema document. Compiling the whole closure once, rather than per part
    /// kind, is what lets an <c>&lt;xsd:import&gt;</c> resolve against another document already in
    /// the set instead of trying to fetch its <c>schemaLocation</c> - there is no
    /// <see cref="XmlResolver"/> to do that fetching with, by design; see <see cref="Load"/>.
    /// </summary>
    private static readonly string[] ResourceNames =
    [
        "sml.xsd",
        "dml-spreadsheetDrawing.xsd",
        "dml-main.xsd",
        "dml-chart.xsd",
        "dml-chartDrawing.xsd",
        "dml-diagram.xsd",
        "dml-lockedCanvas.xsd",
        "dml-picture.xsd",
        "shared-commonSimpleTypes.xsd",
        "shared-relationshipReference.xsd",
        "shared-documentPropertiesExtended.xsd",
        "shared-documentPropertiesCustom.xsd",
        "shared-documentPropertiesVariantTypes.xsd",
    ];

    private static readonly Lazy<XmlSchemaSet> Instance = new(Load);

    /// <summary>The compiled set every schema-validated part is checked against.</summary>
    internal static XmlSchemaSet Set => Instance.Value;

    private static XmlSchemaSet Load()
    {
        XmlSchemaSet schemas = new()
        {
            // The set is meant to be self-contained: every namespace a schema here imports has
            // its own document in ResourceNames, added below, so <xsd:import>/<xsd:include> only
            // ever needs to resolve against another schema already in the set. No resolver means
            // an import this closure does not actually need - none are expected - is silently
            // left unresolved instead of reaching for the network or the filesystem.
            XmlResolver = null,
        };

        foreach (string resourceName in ResourceNames)
        {
            using Stream stream = OpenResource(resourceName);
            using XmlReader reader = XmlReader.Create(stream);
            schemas.Add(targetNamespace: null, reader);
        }

        schemas.Compile();
        return schemas;
    }

    private static Stream OpenResource(string fileName) =>
        Assembly
            .GetExecutingAssembly()
            .GetManifestResourceStream($"XlsxSharp.Excel.IO.Schemas.{fileName}")
        ?? throw new InvalidOperationException(
            $"Embedded schema resource '{fileName}' was not found."
        );
}

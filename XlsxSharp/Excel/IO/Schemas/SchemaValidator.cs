using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO.Schemas;

/// <summary>
/// Validates a saved package against the OOXML schemas, for <see cref="SaveOptions.ValidatePackage"/>.
/// </summary>
/// <remarks>
/// Not every part kind is covered - see the mapping below and <c>PROVENANCE.md</c> next to the
/// embedded schemas for which parts are skipped and why. A part whose content type has no entry
/// there is not validated at all, rather than guessed at: an unmapped content type is either one
/// of those deliberately skipped kinds, or a foreign part this validator has no schema for in the
/// first place, and either way a wrong guess would be worse than no answer.
/// </remarks>
internal static class SchemaValidator
{
    private static readonly XNamespace SmlNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace DrawingMainNs =
        "http://schemas.openxmlformats.org/drawingml/2006/main";

    private static readonly XNamespace SpreadsheetDrawingNs =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private static readonly XNamespace ExtendedPropertiesNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static readonly XNamespace CustomPropertiesNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

    /// <summary>
    /// The root element schema validation expects for a part of the given content type. The
    /// four workbook content types (workbook/template/macro-enabled variants) all carry the exact
    /// same <c>&lt;workbook&gt;</c> content and so map to the same root.
    /// </summary>
    private static readonly Dictionary<string, XName> PartRootElements = new()
    {
        [OoxmlPartTypes.Workbook.ContentType] = SmlNs + "workbook",
        [OoxmlPartTypes.MacroEnabledWorkbook.ContentType] = SmlNs + "workbook",
        [OoxmlPartTypes.WorkbookTemplate.ContentType] = SmlNs + "workbook",
        [OoxmlPartTypes.MacroEnabledWorkbookTemplate.ContentType] = SmlNs + "workbook",
        [OoxmlPartTypes.Worksheet.ContentType] = SmlNs + "worksheet",
        [OoxmlPartTypes.Chartsheet.ContentType] = SmlNs + "chartsheet",
        [OoxmlPartTypes.Styles.ContentType] = SmlNs + "styleSheet",
        [OoxmlPartTypes.SharedStringTable.ContentType] = SmlNs + "sst",
        [OoxmlPartTypes.CalculationChain.ContentType] = SmlNs + "calcChain",
        [OoxmlPartTypes.Theme.ContentType] = DrawingMainNs + "theme",
        [OoxmlPartTypes.PivotTable.ContentType] = SmlNs + "pivotTableDefinition",
        [OoxmlPartTypes.PivotCacheDefinition.ContentType] = SmlNs + "pivotCacheDefinition",
        [OoxmlPartTypes.PivotCacheRecords.ContentType] = SmlNs + "pivotCacheRecords",
        [OoxmlPartTypes.Table.ContentType] = SmlNs + "table",
        [OoxmlPartTypes.Comments.ContentType] = SmlNs + "comments",
        [OoxmlPartTypes.Drawing.ContentType] = SpreadsheetDrawingNs + "wsDr",
        [OoxmlPartTypes.ExtendedFileProperties.ContentType] = ExtendedPropertiesNs + "Properties",
        [OoxmlPartTypes.CustomFileProperties.ContentType] = CustomPropertiesNs + "Properties",
    };

    /// <summary>
    /// Validates every schema-mapped part of <paramref name="package"/>. Returns one message per
    /// problem found, empty when nothing was.
    /// </summary>
    internal static IReadOnlyList<string> Validate(OpcPackage package)
    {
        List<string> errors = [];

        foreach (OpcPart part in package.Parts)
        {
            if (PartRootElements.TryGetValue(part.ContentType, out XName? expectedRoot))
            {
                ValidatePart(part, expectedRoot, errors);
            }
        }

        return errors;
    }

    private static void ValidatePart(OpcPart part, XName expectedRoot, List<string> errors)
    {
        XDocument document;
        using (Stream stream = part.GetReadStream())
        {
            document = XDocument.Load(stream, System.Xml.Linq.LoadOptions.SetLineInfo);
        }

        if (document.Root is null)
        {
            return;
        }

        if (document.Root.Name != expectedRoot)
        {
            errors.Add(
                $"Part {part.Name}: expected root element '{expectedRoot}', found '{document.Root.Name}'."
            );
            return;
        }

        MarkupCompatibility.Strip(document.Root);

        XmlReaderSettings settings = new()
        {
            ValidationType = ValidationType.Schema,
            Schemas = OoxmlSchemas.Set,
        };
        settings.ValidationEventHandler += (_, e) =>
            errors.Add(
                $"Part {part.Name}, line {e.Exception.LineNumber}, position {e.Exception.LinePosition}: {e.Message}"
            );

        using XmlReader reader = XmlReader.Create(document.CreateReader(), settings);
        while (reader.Read()) { }
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

/// <summary>
/// Checks every entry of <see cref="OoxmlPartTypes"/> against the SDK part class it replaces, by
/// letting the SDK create a part of each kind and asking it for its content type and relationship
/// type. This covers the kinds a test workbook does not happen to contain.
/// </summary>
public class OoxmlPartTypeSdkParityTests
{
    [Test]
    public void WorkbookMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.Workbook,
            SpreadsheetDocumentType.Workbook,
            d => d.WorkbookPart!
        );

    [Test]
    public void MacroEnabledWorkbookMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.MacroEnabledWorkbook,
            SpreadsheetDocumentType.MacroEnabledWorkbook,
            d => d.WorkbookPart!
        );

    [Test]
    public void WorkbookTemplateMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.WorkbookTemplate,
            SpreadsheetDocumentType.Template,
            d => d.WorkbookPart!
        );

    [Test]
    public void MacroEnabledWorkbookTemplateMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.MacroEnabledWorkbookTemplate,
            SpreadsheetDocumentType.MacroEnabledTemplate,
            d => d.WorkbookPart!
        );

    [Test]
    public void WorksheetMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Worksheet, AddToWorkbook<WorksheetPart>());

    [Test]
    public void ChartsheetMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Chartsheet, AddToWorkbook<ChartsheetPart>());

    [Test]
    public void StylesMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Styles, AddToWorkbook<WorkbookStylesPart>());

    [Test]
    public void SharedStringTableMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.SharedStringTable, AddToWorkbook<SharedStringTablePart>());

    [Test]
    public void CalculationChainMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.CalculationChain, AddToWorkbook<CalculationChainPart>());

    [Test]
    public void ThemeMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Theme, AddToWorkbook<ThemePart>());

    [Test]
    public void PivotCacheDefinitionMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.PivotCacheDefinition,
            AddToWorkbook<PivotTableCacheDefinitionPart>()
        );

    [Test]
    public void PivotCacheRecordsMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.PivotCacheRecords,
            AddTo<PivotTableCacheRecordsPart>(d =>
                d.WorkbookPart!.AddNewPart<PivotTableCacheDefinitionPart>()
            )
        );

    [Test]
    public void PivotTableMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.PivotTable, AddToWorksheet<PivotTablePart>());

    [Test]
    public void TableMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Table, AddToWorksheet<TableDefinitionPart>());

    [Test]
    public void CommentsMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Comments, AddToWorksheet<WorksheetCommentsPart>());

    [Test]
    public void DrawingMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.Drawing, AddToWorksheet<DrawingsPart>());

    [Test]
    public void VmlDrawingMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.VmlDrawing, AddToWorksheet<VmlDrawingPart>());

    [Test]
    public void ExtendedFilePropertiesMatchesTheSdk() =>
        AssertParity(
            OoxmlPartTypes.ExtendedFileProperties,
            AddTo<ExtendedFilePropertiesPart>(d => d)
        );

    [Test]
    public void CustomFilePropertiesMatchesTheSdk() =>
        AssertParity(OoxmlPartTypes.CustomFileProperties, AddTo<CustomFilePropertiesPart>(d => d));

    private static (string ContentType, string RelationshipType) AddToWorkbook<TPart>()
        where TPart : OpenXmlPart, IFixedContentTypePart => AddTo<TPart>(d => d.WorkbookPart!);

    private static (string ContentType, string RelationshipType) AddToWorksheet<TPart>()
        where TPart : OpenXmlPart, IFixedContentTypePart =>
        AddTo<TPart>(d => d.WorkbookPart!.AddNewPart<WorksheetPart>());

    /// <summary>
    /// Creates a package with the SDK, adds a part of the kind under test where
    /// <paramref name="container"/> says, and reports what the SDK declared for it.
    /// </summary>
    private static (string ContentType, string RelationshipType) AddTo<TPart>(
        Func<SpreadsheetDocument, OpenXmlPartContainer> container
    )
        where TPart : OpenXmlPart, IFixedContentTypePart
    {
        using MemoryStream stream = new();
        using SpreadsheetDocument document = SpreadsheetDocument.Create(
            stream,
            SpreadsheetDocumentType.Workbook
        );

        document.AddWorkbookPart();
        TPart part = container(document).AddNewPart<TPart>();
        return (part.ContentType, part.RelationshipType);
    }

    private static void AssertParity(
        OoxmlPartType partType,
        SpreadsheetDocumentType documentType,
        Func<SpreadsheetDocument, OpenXmlPart> select
    )
    {
        using MemoryStream stream = new();
        using SpreadsheetDocument document = SpreadsheetDocument.Create(stream, documentType);
        document.AddWorkbookPart();

        OpenXmlPart part = select(document);
        AssertParity(partType, (part.ContentType, part.RelationshipType));
    }

    private static void AssertParity(
        OoxmlPartType partType,
        (string ContentType, string RelationshipType) sdk
    )
    {
        ClassicAssert.AreEqual(sdk.ContentType, partType.ContentType, "content type");
        ClassicAssert.AreEqual(
            sdk.RelationshipType,
            partType.RelationshipType,
            "relationship type"
        );
    }
}

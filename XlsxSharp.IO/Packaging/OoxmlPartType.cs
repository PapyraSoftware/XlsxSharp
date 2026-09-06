namespace XlsxSharp.IO.Packaging;

/// <summary>
/// One kind of part in a SpreadsheetML package: the relationship type that points at it, the
/// content type it is declared with, and where a new one is put.
/// </summary>
/// <param name="RelationshipType">The relationship type pointing at parts of this kind.</param>
/// <param name="ContentType">The content type declared for parts of this kind.</param>
/// <param name="PathTemplate">
/// Where a new part goes. A template containing <c>{0}</c> is numbered, so that several parts of
/// the kind can live side by side (<c>/xl/worksheets/sheet1.xml</c>); one without it names a part
/// the package has at most once (<c>/xl/styles.xml</c>).
/// </param>
public sealed record OoxmlPartType(string RelationshipType, string ContentType, string PathTemplate)
{
    /// <summary>Whether a package can hold more than one part of this kind.</summary>
    public bool IsNumbered => this.PathTemplate.Contains("{0}", StringComparison.Ordinal);
}

/// <summary>
/// The part kinds of a SpreadsheetML package, as ECMA-376 Part 1 defines them. This is the table
/// that replaces the SDK's typed part classes: XlsxSharp only ever asked those for their
/// relationship type, their content type and where to put a new one.
/// </summary>
public static class OoxmlPartTypes
{
    private const string OfficeRel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    private const string SpreadsheetType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.";

    /// <summary>
    /// A hyperlink is a relationship with no part of its own - a cell's hyperlink points outside
    /// the package as often as it points at one of its parts - so it does not fit
    /// <see cref="OoxmlPartType"/>, which only carries the kinds of parts a package holds.
    /// </summary>
    public const string HyperlinkRelationshipType = OfficeRel + "hyperlink";

    /// <summary>The workbook part of an <c>.xlsx</c>, i.e. the package's office document.</summary>
    public static OoxmlPartType Workbook { get; } =
        new(OfficeRel + "officeDocument", SpreadsheetType + "sheet.main+xml", "/xl/workbook.xml");

    /// <summary>The workbook part of an <c>.xlsm</c>, which differs only in its content type.</summary>
    public static OoxmlPartType MacroEnabledWorkbook { get; } =
        new(
            OfficeRel + "officeDocument",
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            "/xl/workbook.xml"
        );

    /// <summary>The workbook part of an <c>.xltx</c> template.</summary>
    public static OoxmlPartType WorkbookTemplate { get; } =
        new(
            OfficeRel + "officeDocument",
            SpreadsheetType + "template.main+xml",
            "/xl/workbook.xml"
        );

    /// <summary>The workbook part of an <c>.xltm</c> macro enabled template.</summary>
    public static OoxmlPartType MacroEnabledWorkbookTemplate { get; } =
        new(
            OfficeRel + "officeDocument",
            "application/vnd.ms-excel.template.macroEnabled.main+xml",
            "/xl/workbook.xml"
        );

    public static OoxmlPartType Worksheet { get; } =
        new(
            OfficeRel + "worksheet",
            SpreadsheetType + "worksheet+xml",
            "/xl/worksheets/sheet{0}.xml"
        );

    public static OoxmlPartType Chartsheet { get; } =
        new(
            OfficeRel + "chartsheet",
            SpreadsheetType + "chartsheet+xml",
            "/xl/chartsheets/sheet{0}.xml"
        );

    public static OoxmlPartType Styles { get; } =
        new(OfficeRel + "styles", SpreadsheetType + "styles+xml", "/xl/styles.xml");

    public static OoxmlPartType SharedStringTable { get; } =
        new(
            OfficeRel + "sharedStrings",
            SpreadsheetType + "sharedStrings+xml",
            "/xl/sharedStrings.xml"
        );

    public static OoxmlPartType CalculationChain { get; } =
        new(OfficeRel + "calcChain", SpreadsheetType + "calcChain+xml", "/xl/calcChain.xml");

    public static OoxmlPartType Theme { get; } =
        new(
            OfficeRel + "theme",
            "application/vnd.openxmlformats-officedocument.theme+xml",
            "/xl/theme/theme{0}.xml"
        );

    public static OoxmlPartType PivotTable { get; } =
        new(
            OfficeRel + "pivotTable",
            SpreadsheetType + "pivotTable+xml",
            "/xl/pivotTables/pivotTable{0}.xml"
        );

    /// <summary>
    /// Not under <c>/xl</c> - the one part kind the SDK itself puts at the package root rather
    /// than alongside the rest of the workbook's own parts.
    /// </summary>
    public static OoxmlPartType PivotCacheDefinition { get; } =
        new(
            OfficeRel + "pivotCacheDefinition",
            SpreadsheetType + "pivotCacheDefinition+xml",
            "/pivotCache/pivotCacheDefinition{0}.xml"
        );

    /// <summary>Not under <c>/xl</c>; see <see cref="PivotCacheDefinition"/>.</summary>
    public static OoxmlPartType PivotCacheRecords { get; } =
        new(
            OfficeRel + "pivotCacheRecords",
            SpreadsheetType + "pivotCacheRecords+xml",
            "/pivotCache/pivotCacheRecords{0}.xml"
        );

    public static OoxmlPartType Table { get; } =
        new(OfficeRel + "table", SpreadsheetType + "table+xml", "/xl/tables/table{0}.xml");

    public static OoxmlPartType Comments { get; } =
        new(OfficeRel + "comments", SpreadsheetType + "comments+xml", "/xl/comments{0}.xml");

    public static OoxmlPartType Drawing { get; } =
        new(
            OfficeRel + "drawing",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            "/xl/drawings/drawing{0}.xml"
        );

    /// <summary>
    /// Legacy VML, which carries the shapes of the comments. It is not XML by content type, even
    /// though its content is. The SDK's own default name is all lower case, unlike every other
    /// part kind's, and unlike them it also leaves its first instance unnumbered - so this
    /// template is never handed to <c>AddPartOfType</c> without an explicit part name computed
    /// the same way the SDK numbers it.
    /// </summary>
    public static OoxmlPartType VmlDrawing { get; } =
        new(
            OfficeRel + "vmlDrawing",
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
            "/xl/drawings/vmldrawing{0}.vml"
        );

    public static OoxmlPartType ExtendedFileProperties { get; } =
        new(
            OfficeRel + "extended-properties",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "/docProps/app.xml"
        );

    public static OoxmlPartType CustomFileProperties { get; } =
        new(
            OfficeRel + "custom-properties",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
            "/docProps/custom.xml"
        );

    /// <summary>
    /// Images are the one kind whose content type depends on the individual part rather than on
    /// the kind, so the content type here is only the fallback for an unknown format.
    /// </summary>
    public static OoxmlPartType Image { get; } =
        new(OfficeRel + "image", "image/png", "/xl/media/image{0}.png");
}

using System.Xml.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Placing the children of <c>xl/worksheets/sheetN.xml</c>, which the schema fixes the order of.
/// </summary>
/// <remarks>
/// The worksheet writer patches the sheet the workbook was loaded from rather than writing a
/// fresh one, so an element it adds has to land where the schema says it goes, between whichever
/// of its neighbours the loaded sheet happens to carry.
/// </remarks>
internal static class WorksheetXml
{
    /// <summary>
    /// The children of <c>worksheet</c>, in the order CT_Worksheet requires.
    /// </summary>
    private static readonly string[] ChildOrder =
    [
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetCalcPr",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "dataConsolidate",
        "customSheetViews",
        "mergeCells",
        "phoneticPr",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "customProperties",
        "cellWatches",
        "ignoredErrors",
        "smartTags",
        "drawing",
        "legacyDrawing",
        "legacyDrawingHF",
        "drawingHF",
        "picture",
        "oleObjects",
        "controls",
        "webPublishItems",
        "tableParts",
        "extLst",
    ];

    /// <summary>
    /// The sheet's only child of that name, added in schema order if it has none yet.
    /// </summary>
    internal static XElement Child(XElement worksheet, string name)
    {
        if (worksheet.Element(SpreadsheetXml.Main + name) is { } existing)
        {
            return existing;
        }

        XElement child = new(SpreadsheetXml.Main + name);
        InsertInOrder(worksheet, name, child);
        return child;
    }

    /// <summary>
    /// Adds a child in schema order without looking for one that is already there, for the
    /// elements a sheet may carry more than one of.
    /// </summary>
    internal static void Insert(XElement worksheet, string name, XElement child) =>
        InsertInOrder(worksheet, name, child);

    private static void InsertInOrder(XElement worksheet, string name, XElement child)
    {
        int rank = Array.IndexOf(ChildOrder, name);
        if (rank < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(name), name, "Not a sheet element.");
        }

        XElement? previous = null;
        foreach (XElement candidate in worksheet.Elements())
        {
            int candidateRank = Array.IndexOf(ChildOrder, candidate.Name.LocalName);
            if (
                candidate.Name.Namespace == SpreadsheetXml.Main
                && candidateRank >= 0
                && candidateRank < rank
            )
            {
                previous = candidate;
            }
        }

        if (previous is null)
        {
            worksheet.AddFirst(child);
        }
        else
        {
            previous.AddAfterSelf(child);
        }
    }

    /// <summary>
    /// An attribute carrying an OOXML boolean, which is written as 1 or 0.
    /// </summary>
    internal static void SetBool(XElement element, string name, bool value) =>
        element.SetAttributeValue(name, value ? "1" : "0");

    /// <summary>
    /// An OOXML boolean that is left off the element when it says what the schema already says.
    /// </summary>
    internal static void SetBoolDefault(
        XElement element,
        string name,
        bool value,
        bool defaultValue
    ) =>
        element.SetAttributeValue(
            name,
            value == defaultValue ? null
                : value ? "1"
                : "0"
        );

    /// <summary>
    /// An attribute that is left off the element when it has no value.
    /// </summary>
    internal static void SetOptional<T>(XElement element, string name, T? value)
        where T : struct =>
        element.SetAttributeValue(name, value is { } present ? present.ToInvariantString() : null);

    internal static void Set<T>(XElement element, string name, T value)
        where T : struct => element.SetAttributeValue(name, value.ToInvariantString());
}

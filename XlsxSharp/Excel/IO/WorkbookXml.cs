using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Reading <c>xl/workbook.xml</c>. The counterpart of <see cref="WorkbookPartWriter"/>, and small
/// enough to be a handful of helpers rather than a reader class of its own: the load path reads
/// the part in several places and threads the element through.
/// </summary>
internal static class WorkbookXml
{
    internal static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static XElement Read(WorkbookPart part)
    {
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream).Root
            ?? throw PartStructureException.ExpectedElementNotFound("workbook");
    }

    /// <summary>
    /// OOXML booleans are written as 1/0 or true/false, and both have to be accepted.
    /// </summary>
    /// <remarks>
    /// A value that cannot be read reads as absent rather than as an error, which is what the
    /// SDK's typed values did: an unparseable attribute left HasValue false and the caller took
    /// its default. Files in the wild rely on it - one of the test workbooks carries
    /// activeTab="-1", and the sheet it does not point at is meant to fall back to the first.
    /// </remarks>
    internal static bool? Bool(XElement? element, string name) =>
        element?.Attribute(name)?.Value switch
        {
            "1" or "true" or "on" or "True" => true,
            "0" or "false" or "off" or "False" => false,
            _ => null,
        };

    /// <inheritdoc cref="Bool"/>
    internal static uint? UInt(XElement? element, string name) =>
        uint.TryParse(
            element?.Attribute(name)?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out uint parsed
        )
            ? parsed
            : null;

    /// <inheritdoc cref="Bool"/>
    internal static int? Int(XElement? element, string name) =>
        int.TryParse(
            element?.Attribute(name)?.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out int parsed
        )
            ? parsed
            : null;

    internal static XLWorksheetVisibility ParseSheetState(string value) =>
        value switch
        {
            "visible" => XLWorksheetVisibility.Visible,
            "hidden" => XLWorksheetVisibility.Hidden,
            "veryHidden" => XLWorksheetVisibility.VeryHidden,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLCalculateMode ParseCalculateMode(string value) =>
        value switch
        {
            "auto" => XLCalculateMode.Auto,
            "autoNoTable" => XLCalculateMode.AutoNoTable,
            "manual" => XLCalculateMode.Manual,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLReferenceStyle ParseReferenceMode(string value) =>
        value switch
        {
            "A1" => XLReferenceStyle.A1,
            "R1C1" => XLReferenceStyle.R1C1,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };
}

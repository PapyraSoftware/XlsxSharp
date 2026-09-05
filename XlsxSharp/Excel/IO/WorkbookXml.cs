using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.IO;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Reading <c>xl/workbook.xml</c>. The counterpart of <see cref="WorkbookPartWriter"/>, and small
/// enough to be a handful of helpers rather than a reader class of its own: the load path reads
/// the part in several places and threads the element through.
/// </summary>
internal static class WorkbookXml
{
    internal static XElement Read(OpcPart part)
    {
        using Stream stream = part.GetReadStream();
        return XDocument.Load(stream).Root
            ?? throw PartStructureException.ExpectedElementNotFound("workbook");
    }

    /// <summary>
    /// Reads the part while the save path still opens the package through the SDK. Goes away once
    /// saving moves onto <see cref="OpcPackage"/> too.
    /// </summary>
    internal static XElement Read(WorkbookPart part)
    {
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream).Root
            ?? throw PartStructureException.ExpectedElementNotFound("workbook");
    }

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

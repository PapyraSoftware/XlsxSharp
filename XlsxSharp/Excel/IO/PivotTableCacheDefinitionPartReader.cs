using System.Globalization;
using System.Xml.Linq;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Reads a <c>pivotCacheDefinition</c> part. The counterpart of
/// <see cref="PivotTableCacheDefinitionPartWriter"/>, and reading the same way it writes.
/// </summary>
internal class PivotTableCacheDefinitionPartReader
{
    private static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static XLPivotCache Load(
        OpcPart workbookPart,
        OpcPart pivotTableCacheDefinitionPart,
        XLWorkbook workbook
    )
    {
        XElement cacheDefinition = ReadRoot(pivotTableCacheDefinitionPart);
        IXLPivotSource pivotSourceReference = ParsePivotSourceReference(
            RequireCacheSource(cacheDefinition)
        );

        XLPivotCache pivotCache = workbook.PivotCachesInternal.Add(pivotSourceReference);

        // A WorkbookCacheRelId that is already set means the pivot source is being reused.
        if (string.IsNullOrWhiteSpace(pivotCache.WorkbookCacheRelId))
        {
            pivotCache.WorkbookCacheRelId = workbookPart.Relationships.GetIdOfTarget(
                pivotTableCacheDefinitionPart.Name
            );
        }

        if (ParseUInt(cacheDefinition.Attribute("missingItemsLimit")) is { } missingItemsLimit)
        {
            pivotCache.ItemsToRetainPerField = missingItemsLimit switch
            {
                0 => XLItemsToRetain.None,
                XlsxSharp.XLHelper.MaxRowNumber => XLItemsToRetain.Max,
                _ => XLItemsToRetain.Automatic,
            };
        }

        if (cacheDefinition.Element(Main + "cacheFields") is { } cacheFields)
        {
            ReadCacheFields(cacheFields, pivotCache);
        }

        pivotCache.SaveSourceData = ParseBool(cacheDefinition.Attribute("saveData")) ?? true;
        return pivotCache;
    }

    /// <summary>
    /// The source a cache definition part describes, for the pivot table reader looking for a
    /// cache that reads from the same place.
    /// </summary>
    internal static IXLPivotSource ReadSource(OpcPart part) =>
        ParsePivotSourceReference(RequireCacheSource(ReadRoot(part)));

    private static XElement ReadRoot(OpcPart part)
    {
        using Stream stream = part.GetReadStream();
        return XDocument.Load(stream).Root
            ?? throw PartStructureException.ExpectedElementNotFound("pivotCacheDefinition");
    }

    private static XElement RequireCacheSource(XElement cacheDefinition) =>
        cacheDefinition.Element(Main + "cacheSource")
        ?? throw PartStructureException.RequiredElementIsMissing("cacheSource");

    internal static IXLPivotSource ParsePivotSourceReference(XElement cacheSource)
    {
        // Cache source has several types. Each has a specific required format. Do not use different
        // combinations, Excel will crash or at least try to repair
        // [worksheet] uses a worksheet source:
        //   * An unnamed range in a sheet: Uses `sheet` and `ref`.
        //   * An table: Uses `name` that contains a name of the table.
        // [external]
        //   * `connectionId` link to external relationships.
        // [consolidation]
        //  * uses consolidation tag and a list of range sets plus optionally
        //    page fields to add a custom report fields that allow user to select
        //    ranges from rangeSet to calculate values.
        // [scenario]
        //  * only type attribute tag is specified, no other value. Likely linked
        //    through cacheField names (e.g. <cacheField name="$A$1 by">).

        // Not all sources are supported, but at least pipe the data through so the load/save works
        string sourceType =
            cacheSource.Attribute("type")?.Value ?? throw PartStructureException.MissingAttribute();

        switch (sourceType)
        {
            case "worksheet":
            {
                XElement sheetSource =
                    cacheSource.Element(Main + "worksheetSource")
                    ?? throw PartStructureException.ExpectedElementNotFound(
                        "'worksheetSource' element is required for type 'worksheet'."
                    );

                string? externalWorkbookRelId = sheetSource.Attribute(Rel + "id")?.Value;

                // If the source is a defined name, it must be a single area reference
                if (sheetSource.Attribute("name")?.Value is { } tableOrName)
                {
                    return externalWorkbookRelId is not null
                        ? new XLPivotSourceExternalWorkbook(externalWorkbookRelId, tableOrName)
                        : new XLPivotSourceReference(tableOrName);
                }

                if (
                    sheetSource.Attribute("sheet")?.Value is { } sheetName
                    && sheetSource.Attribute("ref")?.Value is { } areaRef
                    && Area.TryParse(areaRef.AsSpan(), out Area sheetArea)
                )
                {
                    SheetArea area = new(sheetName, sheetArea);
                    return externalWorkbookRelId is not null
                        ? new XLPivotSourceExternalWorkbook(externalWorkbookRelId, area)
                        : new XLPivotSourceReference(area);
                }

                throw PartStructureException.IncorrectElementFormat("worksheetSource");
            }

            case "external":
            {
                if (ParseUInt(cacheSource.Attribute("connectionId")) is not { } connectionId)
                {
                    throw PartStructureException.MissingAttribute("connectionId");
                }

                return new XLPivotSourceConnection(connectionId);
            }

            case "consolidation":
                return ParseConsolidation(
                    cacheSource.Element(Main + "consolidation")
                        ?? throw PartStructureException.ExpectedElementNotFound("consolidation")
                );

            case "scenario":
                return new XLPivotSourceScenario();

            default:
                throw PartStructureException.InvalidAttributeValue(sourceType);
        }
    }

    private static XLPivotSourceConsolidation ParseConsolidation(XElement consolidation)
    {
        bool autoPage = ParseBool(consolidation.Attribute("autoPage")) ?? true;

        List<XLPivotCacheSourceConsolidationPage> xlPages = [];
        if (consolidation.Element(Main + "pages") is { } pages)
        {
            // There is 1..4 pages
            foreach (XElement page in pages.Elements(Main + "page"))
            {
                List<string> xlPageItems = [];
                foreach (XElement pageItem in page.Elements(Main + "pageItem"))
                {
                    xlPageItems.Add(
                        pageItem.Attribute("name")?.Value
                            ?? throw PartStructureException.MissingAttribute()
                    );
                }

                xlPages.Add(new XLPivotCacheSourceConsolidationPage(xlPageItems));
            }
        }

        if (consolidation.Element(Main + "rangeSets") is not { } rangeSets)
        {
            throw PartStructureException.RequiredElementIsMissing("rangeSets");
        }

        List<XLPivotCacheSourceConsolidationRangeSet> xlRangeSets =
        [
            .. rangeSets.Elements(Main + "rangeSet").Select(r => GetRangeSet(r, xlPages)),
        ];

        if (xlRangeSets.Count < 1)
        {
            throw PartStructureException.IncorrectElementsCount();
        }

        return new XLPivotSourceConsolidation
        {
            AutoPage = autoPage,
            Pages = xlPages,
            RangeSets = xlRangeSets,
        };
    }

    private static XLPivotCacheSourceConsolidationRangeSet GetRangeSet(
        XElement rangeSet,
        List<XLPivotCacheSourceConsolidationPage> xlPages
    )
    {
        uint?[] pageIndexes =
        [
            ParseUInt(rangeSet.Attribute("i1")),
            ParseUInt(rangeSet.Attribute("i2")),
            ParseUInt(rangeSet.Attribute("i3")),
            ParseUInt(rangeSet.Attribute("i4")),
        ];

        // Validate that supplied indexes reference existing page and page items
        for (int i = 0; i < pageIndexes.Length; ++i)
        {
            uint? pageIndex = pageIndexes[i];

            // If there is a page and rangeSet doesn't define index to the page, it is displayed as blank
            if (pageIndex is null)
            {
                continue;
            }

            // Range set points to a non-existent page filter
            if (i >= xlPages.Count)
            {
                throw PartStructureException.InvalidAttributeValue();
            }

            // Range set points to a non-existent item in a page filter
            XLPivotCacheSourceConsolidationPage pageFilter = xlPages[i];
            if (pageIndex.Value >= pageFilter.PageItems.Count)
            {
                throw PartStructureException.InvalidAttributeValue();
            }
        }

        string? relId = rangeSet.Attribute(Rel + "id")?.Value;

        if (rangeSet.Attribute("name")?.Value is { } tableOrName)
        {
            return new XLPivotCacheSourceConsolidationRangeSet
            {
                Indexes = pageIndexes,
                RelId = relId,
                TableOrName = tableOrName,
            };
        }

        if (
            rangeSet.Attribute("sheet")?.Value is { } sheet
            && rangeSet.Attribute("ref")?.Value is { } reference
            && Area.TryParse(reference.AsSpan(), out Area area)
        )
        {
            return new XLPivotCacheSourceConsolidationRangeSet
            {
                Indexes = pageIndexes,
                RelId = relId,
                Area = new SheetArea(sheet, area),
            };
        }

        throw PartStructureException.IncorrectElementFormat("rangeSet");
    }

    private static void ReadCacheFields(XElement cacheFields, XLPivotCache pivotCache)
    {
        foreach (XElement cacheField in cacheFields.Elements(Main + "cacheField"))
        {
            if (cacheField.Attribute("name")?.Value is not { } fieldName)
            {
                throw PartStructureException.MissingAttribute();
            }

            if (pivotCache.ContainsField(fieldName))
            {
                // We don't allow duplicate field names... but what do we do if we find one? Let's just skip it.
                continue;
            }

            XElement? sharedItems = cacheField.Element(Main + "sharedItems");
            XLPivotCacheValuesStats fieldStats = ReadCacheFieldStats(sharedItems);
            XLPivotCacheSharedItems fieldSharedItems = sharedItems is not null
                ? ReadSharedItems(sharedItems)
                : new XLPivotCacheSharedItems();

            pivotCache.AddCachedField(
                fieldName,
                new XLPivotCacheValues(fieldSharedItems, fieldStats)
            );
        }
    }

    private static XLPivotCacheValuesStats ReadCacheFieldStats(XElement? sharedItems)
    {
        // Various statistics about the records of the field, not just shared items. The
        // containsMixedTypes, containsNonDate and containsSemiMixedTypes are derived from these.
        return new XLPivotCacheValuesStats(
            ParseBool(sharedItems?.Attribute("containsBlank")) ?? false,
            ParseBool(sharedItems?.Attribute("containsNumber")) ?? false,
            ParseBool(sharedItems?.Attribute("containsInteger")) ?? false,
            ParseDouble(sharedItems?.Attribute("minValue")),
            ParseDouble(sharedItems?.Attribute("maxValue")),
            ParseBool(sharedItems?.Attribute("containsString")) ?? true,
            ParseBool(sharedItems?.Attribute("longText")) ?? false,
            ParseBool(sharedItems?.Attribute("containsDate")) ?? false,
            ParseDateTime(sharedItems?.Attribute("minDate")),
            ParseDateTime(sharedItems?.Attribute("maxDate"))
        );
    }

    private static XLPivotCacheSharedItems ReadSharedItems(XElement fieldSharedItems)
    {
        XLPivotCacheSharedItems sharedItems = new();

        foreach (XElement item in fieldSharedItems.Elements())
        {
            // Shared items can't contain element of type index (`x`), because index references
            // shared items. That is the main reason for the duplication with reading records.
            switch (item.Name.LocalName)
            {
                case "m":
                    sharedItems.AddMissing();
                    break;

                case "n":
                    sharedItems.AddNumber(
                        ParseDouble(RequireValue(item))
                            ?? throw PartStructureException.InvalidAttributeFormat()
                    );

                    break;

                case "b":
                    sharedItems.AddBoolean(
                        ParseBool(RequireValue(item))
                            ?? throw PartStructureException.InvalidAttributeFormat()
                    );

                    break;

                case "e":
                    if (!XLErrorParser.TryParseError(RequireValue(item).Value, out XLError error))
                    {
                        throw PartStructureException.InvalidAttributeFormat();
                    }

                    sharedItems.AddError(error);
                    break;

                case "s":
                    sharedItems.AddString(RequireValue(item).Value);
                    break;

                case "d":
                    sharedItems.AddDateTime(
                        ParseDateTime(RequireValue(item))
                            ?? throw PartStructureException.InvalidAttributeFormat()
                    );

                    break;

                default:
                    throw PartStructureException.ExpectedElementNotFound();
            }
        }

        return sharedItems;
    }

    private static XAttribute RequireValue(XElement item) =>
        item.Attribute("v") ?? throw PartStructureException.MissingAttribute();

    /// <summary>
    /// OOXML booleans are written as 1/0 or true/false, and both have to be accepted on the way
    /// in even though only the short form is written back.
    /// </summary>
    private static bool? ParseBool(XAttribute? attribute) =>
        attribute?.Value switch
        {
            null => null,
            "1" or "true" or "on" or "True" => true,
            "0" or "false" or "off" or "False" => false,
            _ => throw PartStructureException.InvalidAttributeFormat(),
        };

    private static uint? ParseUInt(XAttribute? attribute) =>
        attribute is null ? null
        : uint.TryParse(
            attribute.Value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out uint value
        )
            ? value
        : throw PartStructureException.InvalidAttributeFormat();

    private static double? ParseDouble(XAttribute? attribute) =>
        attribute is null ? null
        : double.TryParse(
            attribute.Value,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out double value
        )
            ? value
        : throw PartStructureException.InvalidAttributeFormat();

    private static DateTime? ParseDateTime(XAttribute? attribute) =>
        attribute is null ? null
        : DateTime.TryParse(
            attribute.Value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out DateTime value
        )
            ? value
        : throw PartStructureException.InvalidAttributeFormat();
}

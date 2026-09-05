#nullable disable

using System.Diagnostics;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writes a <c>pivotCacheDefinition</c> part.
/// </summary>
/// <remarks>
/// The part is patched, not rewritten. A cache definition that came from Excel carries plenty
/// XlsxSharp does not model - <c>extLst</c>, who refreshed it and when, the revision uid - and
/// all of it has to survive a load and save, so only the source and the fields are replaced.
/// </remarks>
internal class PivotTableCacheDefinitionPartWriter
{
    private static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly XNamespace Rel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Child order on the root, from the schema (ECMA-376 Part 1 §18.10.1.67). Element order is
    /// part of what the reference workbooks compare, unlike attribute order.
    /// </summary>
    private static readonly string[] ElementOrder =
    [
        "cacheSource",
        "cacheFields",
        "cacheHierarchies",
        "kpis",
        "tupleCache",
        "calculatedItems",
        "calculatedMembers",
        "dimensions",
        "measureGroups",
        "maps",
        "extLst",
    ];

    internal static void GenerateContent(
        PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart,
        XLPivotCache pivotCache,
        SaveContext context
    )
    {
        (XDocument document, bool isNew) = ReadExisting(pivotTableCacheDefinitionPart);
        XElement root = document.Root;

        if (isNew)
        {
            root.SetAttributeValue(Rel + "id", "rId1");
        }

        // The three versions only ever go up: a workbook written by a newer Excel must not be
        // told it was created by an older one.
        SetVersion(root, "createdVersion", XLConstants.PivotTable.CreatedVersion);
        SetVersion(root, "refreshedVersion", XLConstants.PivotTable.RefreshedVersion);
        SetVersion(root, "minRefreshableVersion", 3);

        root.SetAttributeValue("saveData", Bool(pivotCache.SaveSourceData));
        root.SetAttributeValue("refreshOnLoad", Bool(true));

        if (pivotCache.ItemsToRetainPerField == XLItemsToRetain.None)
        {
            root.SetAttributeValue("missingItemsLimit", "0");
        }
        else if (pivotCache.ItemsToRetainPerField == XLItemsToRetain.Max)
        {
            root.SetAttributeValue(
                "missingItemsLimit",
                XlsxSharp.XLHelper.MaxRowNumber.ToInvariantString()
            );
        }

        SetElement(root, "cacheSource", BuildCacheSource(pivotCache));

        // Only the fields themselves are rebuilt. The cacheFields element that a loaded part
        // brought along keeps its own attributes - Excel writes a count there, and the previous
        // writer kept it because it emptied the element instead of replacing it.
        SetElement(root, "cacheFields", BuildCacheFields(pivotCache), keepAttributes: true);

        using Stream partStream = pivotTableCacheDefinitionPart.GetStream(FileMode.Create);
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );

        document.Save(xml);
    }

    /// <summary>
    /// The part's document with the namespaces the writer needs declared on the root, or a fresh
    /// one when the part is empty.
    /// </summary>
    /// <remarks>
    /// The prefixes matter: the reference workbooks record which declarations sit on the root,
    /// and a default namespace where a prefixed one is expected counts as a difference even
    /// though the elements are the same. A part loaded from Excel declares the main namespace as
    /// the default one, so that declaration is dropped and the prefixed pair put in its place,
    /// which is what the SDK did when it re-serialised the part.
    /// </remarks>
    private static (XDocument Document, bool IsNew) ReadExisting(PivotTableCacheDefinitionPart part)
    {
        XElement loaded = null;
        bool standalone = false;

        using (Stream stream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read))
        {
            if (stream.Length > 0)
            {
                try
                {
                    XDocument existing = XDocument.Load(stream);
                    loaded = existing.Root;
                    standalone = string.Equals(
                        existing.Declaration?.Standalone,
                        "yes",
                        StringComparison.OrdinalIgnoreCase
                    );
                }
                catch (XmlException)
                {
                    // A cache definition we cannot read is one we are about to replace anyway.
                }
            }
        }

        if (loaded is null)
        {
            return (
                new XDocument(
                    new XElement(
                        Main + "pivotCacheDefinition",
                        new XAttribute(XNamespace.Xmlns + "r", Rel.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "x", Main.NamespaceName)
                    )
                ),
                true
            );
        }

        XElement root = new(Main + "pivotCacheDefinition");

        // Carry over every declaration except the default one, then make sure the two the writer
        // itself needs are there.
        foreach (XAttribute attribute in loaded.Attributes())
        {
            if (attribute.IsNamespaceDeclaration)
            {
                if (attribute.Name.LocalName != "xmlns")
                {
                    root.Add(new XAttribute(attribute));
                }

                continue;
            }

            root.Add(new XAttribute(attribute));
        }

        EnsureDeclaration(root, "r", Rel);
        EnsureDeclaration(root, "x", Main);

        foreach (XElement child in loaded.Elements())
        {
            XElement copy = new(child);
            copy.DescendantsAndSelf()
                .Attributes()
                .Where(a => a.IsNamespaceDeclaration && a.Name.LocalName == "xmlns")
                .ToList()
                .ForEach(a => a.Remove());

            root.Add(copy);
        }

        HoistDeclarations(root);

        return (
            standalone
                ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root)
                : new XDocument(root),
            false
        );
    }

    /// <summary>
    /// Copies the namespace declarations of the descendants up onto the root, leaving them where
    /// they are as well.
    /// </summary>
    /// <remarks>
    /// This is what the SDK did when it re-serialised a part: a prefix that a workbook from Excel
    /// declares only where it is used, on an <c>ext</c> for instance, comes back out declared on
    /// the root too. Without it the root is one declaration short of what the reference workbooks
    /// record.
    /// </remarks>
    private static void HoistDeclarations(XElement root)
    {
        foreach (XElement descendant in root.Descendants())
        {
            foreach (XAttribute attribute in descendant.Attributes().ToList())
            {
                if (!attribute.IsNamespaceDeclaration || attribute.Name.LocalName == "xmlns")
                {
                    continue;
                }

                bool taken = root.Attributes()
                    .Any(a => a.IsNamespaceDeclaration && a.Name == attribute.Name);

                if (!taken)
                {
                    root.Add(new XAttribute(attribute));
                }
            }
        }
    }

    private static void EnsureDeclaration(XElement root, string prefix, XNamespace ns)
    {
        bool declared = root.Attributes()
            .Any(a => a.IsNamespaceDeclaration && a.Value == ns.NamespaceName);

        if (!declared)
        {
            root.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
        }
    }

    private static void SetVersion(XElement root, string name, byte version)
    {
        XAttribute existing = root.Attribute(name);
        byte value =
            existing is not null && byte.TryParse(existing.Value, out byte loaded)
                ? Math.Max(version, loaded)
                : version;

        root.SetAttributeValue(name, value.ToInvariantString());
    }

    /// <summary>
    /// Replaces a child the writer owns, keeping its position when the part already had it and
    /// otherwise putting it where <see cref="ElementOrder"/> says.
    /// </summary>
    private static void SetElement(
        XElement root,
        string localName,
        XElement replacement,
        bool keepAttributes = false
    )
    {
        XElement existing = root.Element(Main + localName);
        if (existing is not null)
        {
            if (keepAttributes)
            {
                foreach (XAttribute attribute in existing.Attributes())
                {
                    replacement.SetAttributeValue(attribute.Name, attribute.Value);
                }
            }

            existing.ReplaceWith(replacement);
            return;
        }

        XElement predecessor = null;
        int position = Array.IndexOf(ElementOrder, localName);
        for (int i = 0; i < position; i++)
        {
            XElement candidate = root.Element(Main + ElementOrder[i]);
            if (candidate is not null)
            {
                predecessor = candidate;
            }
        }

        if (predecessor is null)
        {
            root.AddFirst(replacement);
        }
        else
        {
            predecessor.AddAfterSelf(replacement);
        }
    }

    private static XElement BuildCacheSource(XLPivotCache pivotCache)
    {
        XElement cacheSource = new(Main + "cacheSource");

        switch (pivotCache.Source)
        {
            case XLPivotSourceReference localSource:
            {
                cacheSource.SetAttributeValue("type", "worksheet");

                // Do not quote a worksheet name with whitespace here, see issue #955.
                XElement worksheetSource = new(Main + "worksheetSource");
                if (localSource.UsesName)
                {
                    worksheetSource.SetAttributeValue("name", localSource.Name);
                }
                else
                {
                    worksheetSource.SetAttributeValue(
                        "ref",
                        localSource.Area.Value.Area.ToString()
                    );

                    worksheetSource.SetAttributeValue("sheet", localSource.Area.Value.Name);
                }

                cacheSource.Add(worksheetSource);
                break;
            }

            case XLPivotSourceExternalWorkbook externalSource:
            {
                cacheSource.SetAttributeValue("type", "worksheet");

                XElement worksheetSource = new(Main + "worksheetSource");
                worksheetSource.SetAttributeValue(Rel + "id", externalSource.RelId);
                if (externalSource.UsesName)
                {
                    worksheetSource.SetAttributeValue("name", externalSource.TableOrName);
                }
                else
                {
                    worksheetSource.SetAttributeValue(
                        "ref",
                        externalSource.Area.Value.Area.ToString()
                    );

                    worksheetSource.SetAttributeValue("sheet", externalSource.Area.Value.Name);
                }

                cacheSource.Add(worksheetSource);
                break;
            }

            case XLPivotSourceConnection connectionSource:
                cacheSource.SetAttributeValue("type", "external");
                cacheSource.SetAttributeValue(
                    "connectionId",
                    connectionSource.ConnectionId.ToInvariantString()
                );

                break;

            case XLPivotSourceConsolidation consolidationSource:
                cacheSource.SetAttributeValue("type", "consolidation");
                cacheSource.Add(BuildConsolidation(consolidationSource));
                break;

            case XLPivotSourceScenario:
                cacheSource.SetAttributeValue("type", "scenario");
                break;

            default:
                throw new UnreachableException();
        }

        return cacheSource;
    }

    private static XElement BuildConsolidation(XLPivotSourceConsolidation source)
    {
        XElement consolidation = new(Main + "consolidation");
        consolidation.SetAttributeValue("autoPage", Bool(source.AutoPage));

        if (source.Pages.Count > 0)
        {
            XElement pages = new(Main + "pages");
            foreach (XLPivotCacheSourceConsolidationPage xlPageFilter in source.Pages)
            {
                XElement page = new(Main + "page");
                foreach (string xlPageItem in xlPageFilter.PageItems)
                {
                    page.Add(new XElement(Main + "pageItem", new XAttribute("name", xlPageItem)));
                }

                pages.Add(page);
            }

            consolidation.Add(pages);
        }

        XElement rangeSets = new(Main + "rangeSets");
        foreach (XLPivotCacheSourceConsolidationRangeSet xlRangeSet in source.RangeSets)
        {
            IReadOnlyList<uint?> indexes = xlRangeSet.Indexes;
            XElement rangeSet = new(Main + "rangeSet");
            SetIndex(rangeSet, "i1", indexes, 0);
            SetIndex(rangeSet, "i2", indexes, 1);
            SetIndex(rangeSet, "i3", indexes, 2);
            SetIndex(rangeSet, "i4", indexes, 3);

            // An unset value has to stay absent rather than become an empty string.
            if (xlRangeSet.RelId is not null)
            {
                rangeSet.SetAttributeValue(Rel + "id", xlRangeSet.RelId);
            }

            if (xlRangeSet.UsesName)
            {
                rangeSet.SetAttributeValue("name", xlRangeSet.TableOrName);
            }
            else
            {
                SheetArea rangeArea = xlRangeSet.Area.Value;
                rangeSet.SetAttributeValue("sheet", rangeArea.Name);
                rangeSet.SetAttributeValue("ref", rangeArea.Area.ToString());
            }

            rangeSets.Add(rangeSet);
        }

        consolidation.Add(rangeSets);
        return consolidation;
    }

    private static void SetIndex(
        XElement rangeSet,
        string name,
        IReadOnlyList<uint?> indexes,
        int position
    )
    {
        uint? value = indexes.Count > position ? indexes[position] : null;
        if (value is not null)
        {
            rangeSet.SetAttributeValue(name, value.Value.ToInvariantString());
        }
    }

    private static XElement BuildCacheFields(XLPivotCache pivotCache)
    {
        XElement cacheFields = new(Main + "cacheFields");

        for (int fieldIdx = 0; fieldIdx < pivotCache.FieldCount; ++fieldIdx)
        {
            XLPivotCacheValues fieldValues = pivotCache.GetFieldValues(fieldIdx);
            XLCellValue[] xlSharedItems =
            [
                .. pivotCache.GetFieldSharedItems(fieldIdx).GetCellValues(),
            ];

            XElement cacheField = new(Main + "cacheField");
            cacheField.SetAttributeValue("name", pivotCache.FieldNames[fieldIdx]);
            cacheField.Add(BuildSharedItems(fieldValues, xlSharedItems));
            cacheFields.Add(cacheField);
        }

        return cacheFields;
    }

    private static XElement BuildSharedItems(
        XLPivotCacheValues fieldValues,
        XLCellValue[] xlSharedItems
    )
    {
        XLPivotCacheValuesStats stats = fieldValues.Stats;
        XElement sharedItems = new(Main + "sharedItems");

        if (fieldValues.SharedCount != 0)
        {
            sharedItems.SetAttributeValue(
                "count",
                checked((uint)xlSharedItems.Length).ToInvariantString()
            );
        }

        // https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shareditems
        // The attributes below are not required or used when there are no items in sharedItems.
        SetOptionalBool(sharedItems, "containsBlank", stats.ContainsBlank, false);
        SetOptionalBool(sharedItems, "containsDate", stats.ContainsDate, false);

        // Blank is not a type in OOXML, it is a value, so it does not count here.
        int typesCount =
            (stats.ContainsNumber ? 1 : 0)
            + (stats.ContainsString ? 1 : 0)
            + (stats.ContainsDate ? 1 : 0);

        // ISO29500: whether this field contains more than one data type.
        // MS-OI29500: Office counts boolean and error as part of the string type.
        SetOptionalBool(sharedItems, "containsMixedTypes", typesCount > 1, false);

        // ISO29500: whether the field contains at least one value that is not a date.
        SetOptionalBool(
            sharedItems,
            "containsNonDate",
            stats.ContainsString || stats.ContainsNumber,
            true
        );

        if (stats.ContainsDate)
        {
            // A field with a date treats its numbers as serial date times. Excel repairs the
            // cache definition if both containsNumber and containsDate are given, so only the
            // date bounds go out.

            // The serial to date conversion is the exception to "1900 is a leap year": values
            // are stored counting from 1899-12-30.
            long? minValueAsDateTime = stats.MinValue is not null
                ? DateTime.FromOADate(stats.MinValue.Value).Ticks
                : null;
            long? maxValueAsDateTime = stats.MaxValue is not null
                ? DateTime.FromOADate(stats.MaxValue.Value).Ticks
                : null;

            long? minDateTicks = Min(stats.MinDate?.Ticks, minValueAsDateTime);
            long? maxDateTicks = Max(stats.MaxDate?.Ticks, maxValueAsDateTime);

            // minDate and maxDate may only be present if at least one child is a d element.
            if (minDateTicks is not null)
            {
                sharedItems.SetAttributeValue("minDate", DateTimeValue(new(minDateTicks.Value)));
            }

            if (maxDateTicks is not null)
            {
                sharedItems.SetAttributeValue("maxDate", DateTimeValue(new(maxDateTicks.Value)));
            }

            static long? Min(long? val1, long? val2) =>
                val1 is null || val2 is null ? val1 ?? val2 : Math.Min(val1.Value, val2.Value);

            static long? Max(long? val1, long? val2) =>
                val1 is null || val2 is null ? val1 ?? val2 : Math.Max(val1.Value, val2.Value);
        }
        else if (stats.ContainsNumber)
        {
            SetOptionalBool(sharedItems, "containsNumber", stats.ContainsNumber, false);

            // containsInteger requires containsNumber, MS-OI29500: Office expects containsNumber
            // to be true when containsInteger is given, and reads containsInteger as "only
            // integers, no non-integer numbers".
            SetOptionalBool(sharedItems, "containsInteger", stats.ContainsInteger, false);

            if (stats.MinValue is not null)
            {
                sharedItems.SetAttributeValue("minValue", stats.MinValue.Value.ToInvariantString());
            }

            if (stats.MaxValue is not null)
            {
                sharedItems.SetAttributeValue("maxValue", stats.MaxValue.Value.ToInvariantString());
            }
        }

        // ISO29500: at least one text value, possibly mixed with other types and blanks.
        // MS-OI29500: Office expects this when the field contains text, blank, boolean or error.
        SetOptionalBool(
            sharedItems,
            "containsSemiMixedTypes",
            stats.ContainsString || stats.ContainsBlank,
            true
        );

        // MS-OI29500: Office counts boolean and error as strings here.
        SetOptionalBool(sharedItems, "containsString", stats.ContainsString, true);
        SetOptionalBool(sharedItems, "longText", stats.LongText, false);

        foreach (XLCellValue value in xlSharedItems)
        {
            sharedItems.Add(BuildSharedItem(value));
        }

        return sharedItems;
    }

    private static XElement BuildSharedItem(XLCellValue value) =>
        value.Type switch
        {
            XLDataType.Blank => new XElement(Main + "m"),
            XLDataType.Boolean => Item("b", Bool(value.GetBoolean())),
            XLDataType.Number => Item("n", value.GetNumber().ToInvariantString()),
            XLDataType.Text => Item("s", value.GetText()),
            XLDataType.Error => Item("e", value.GetError().ToDisplayString()),
            XLDataType.DateTime => Item("d", DateTimeValue(value.GetDateTime())),
            XLDataType.TimeSpan => Item(
                "d",
                DateTimeValue(DateTime.FromOADate(value.GetUnifiedNumber()))
            ),
            _ => throw new InvalidOperationException(),
        };

    private static XElement Item(string localName, string value) =>
        new(Main + localName, new XAttribute("v", value));

    /// <summary>
    /// Writes the attribute only when it differs from the value a reader would assume, which is
    /// what the SDK's optional boolean values did.
    /// </summary>
    private static void SetOptionalBool(
        XElement element,
        string name,
        bool value,
        bool defaultValue
    )
    {
        if (value != defaultValue)
        {
            element.SetAttributeValue(name, Bool(value));
        }
    }

    private static string Bool(bool value) => value ? "1" : "0";

    private static string DateTimeValue(DateTime value) =>
        value.ToString("yyyy-MM-ddTHH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
}

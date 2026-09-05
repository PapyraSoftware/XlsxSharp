using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Reads a <c>pivotTableDefinition</c> part.
/// </summary>
/// <remarks>
/// The attribute names are not always what the workbook model calls them, and a name that does
/// not match reads as absent and silently takes the default. They were taken from the SDK rather
/// than from the schema; <see cref="PivotXmlEnums"/> does the same for the enumerations.
/// </remarks>
internal class PivotTableDefinitionPartReader
{
    private static readonly XNamespace Main =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>The 2009 extension namespace, which carries the few x14 attributes read here.</summary>
    private static readonly XNamespace X14 =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    /// <summary>
    /// A field displayed as <c>∑Values</c> in a pivot table that contains names of all aggregation
    /// function in value fields collection. Also commonly called 'data' field.
    /// </summary>
    private const int ValuesFieldIndex = -2;

    internal static void Load(
        WorkbookPart workbookPart,
        PivotTablePart pivotTablePart,
        WorksheetPart worksheetPart,
        XLWorksheet ws,
        LoadContext context
    )
    {
        XLWorkbook workbook = ws.Workbook;
        // Without a cache there is nothing to read the pivot table against. The previous reader
        // would have thrown a NullReferenceException a line later.
        PivotTableCacheDefinitionPart cache =
            pivotTablePart.PivotTableCacheDefinitionPart
            ?? throw PartStructureException.ExpectedElementNotFound(
                "the pivot table has no cache definition part"
            );
        string cacheDefinitionRelId = workbookPart.GetIdOfPart(cache);

        XLPivotCache? pivotSource = workbook.PivotCachesInternal.FirstOrDefault<XLPivotCache>(ps =>
            ps.WorkbookCacheRelId == cacheDefinitionRelId
        );

        if (pivotSource is null)
        {
            // If it's missing, find a 'similar' pivot cache, i.e. one that's based on the same
            // source range/table. Reading the part again is cheap next to the alternative, which
            // is keeping the cache definition's DOM alive just for this fallback.
            IXLPivotSource cacheSource = PivotTableCacheDefinitionPartReader.ReadSource(cache);
            pivotSource = workbook.PivotCachesInternal.FirstOrDefault<XLPivotCache>(ps =>
                ps.Source.Equals(cacheSource)
            );
        }

        XElement? pivotTableDefinition = ReadRoot(pivotTablePart);

        XLCell? target = ws.FirstCell();
        string? locationReference = pivotTableDefinition
            ?.Element(Main + "location")
            ?.Attribute("ref")
            ?.Value;

        if (locationReference is not null && ws.Range(locationReference) is { } locationRange)
        {
            locationRange.Clear();
            target = (XLCell)locationRange.FirstCell();
        }

        if (target is not null && pivotSource is not null && pivotTableDefinition is not null)
        {
            XLPivotTable pt = LoadPivotTableDefinition(
                pivotTableDefinition,
                ws,
                pivotSource,
                context
            );

            ws.PivotTables.Add(pt);

            pt.RelId = worksheetPart.GetIdOfPart(pivotTablePart);
            pt.CacheDefinitionRelId = pivotTablePart.GetIdOfPart(cache);
        }
    }

    private static XElement? ReadRoot(PivotTablePart part)
    {
        using Stream stream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
        return stream.Length > 0 ? XDocument.Load(stream).Root : null;
    }

    private static XLPivotTable LoadPivotTableDefinition(
        XElement pivotTable,
        XLWorksheet sheet,
        XLPivotCache cache,
        LoadContext context
    )
    {
        XLWorkbookStyles styles = sheet.Workbook.Styles;

        // Load base attributes
        XLPivotTable xlPivotTable = LoadPivotTableAttributes(pivotTable, sheet, cache);

        // Load location
        XElement location =
            pivotTable.Element(Main + "location")
            ?? throw PartStructureException.ExpectedElementNotFound();

        xlPivotTable.Area = Area.Parse(RequiredString(location, "ref"));
        xlPivotTable.FirstHeaderRow = RequiredUInt(location, "firstHeaderRow");
        xlPivotTable.FirstDataRow = RequiredUInt(location, "firstDataRow");
        xlPivotTable.FirstDataCol = RequiredUInt(location, "firstDataCol");

        // Skip `rowPageCount` and `colPageCount`, because they are derived from filterAreaOrder, filterFieldsPageWrap and pageField count

        // Load pivot fields
        foreach (XElement pivotField in Children(pivotTable, "pivotFields", "pivotField"))
        {
            xlPivotTable.AddField(LoadPivotField(pivotField, xlPivotTable, styles));
        }

        // Load row axis fields and items
        LoadAxisFields(pivotTable.Element(Main + "rowFields"), xlPivotTable.RowAxis, xlPivotTable);
        LoadAxisItems(pivotTable.Element(Main + "rowItems"), xlPivotTable.RowAxis);

        // Load column axis fields and items
        LoadAxisFields(
            pivotTable.Element(Main + "colFields"),
            xlPivotTable.ColumnAxis,
            xlPivotTable
        );
        LoadAxisItems(pivotTable.Element(Main + "colItems"), xlPivotTable.ColumnAxis);

        // Load page fields, i.e. the filters region.
        foreach (XElement pageField in Children(pivotTable, "pageFields", "pageField"))
        {
            XLPivotPageField xlPageField = new(RequiredInt(pageField, "fld"))
            {
                ItemIndex = checked((int?)OptionalUInt(pageField, "item")),
                HierarchyIndex = OptionalInt(pageField, "hier"),
                HierarchyUniqueName = pageField.Attribute("name")?.Value,
                HierarchyDisplayName = pageField.Attribute("cap")?.Value,
            };

            xlPivotTable.Filters.AddField(xlPageField);
        }

        // Load data fields.
        foreach (XElement dataField in Children(pivotTable, "dataFields", "dataField"))
        {
            int? numberFormatId = checked((int?)OptionalUInt(dataField, "numFmtId"));
            XLPivotDataField xlDataField = new(
                xlPivotTable,
                checked((int)RequiredUInt(dataField, "fld"))
            )
            {
                DataFieldName = dataField.Attribute("name")?.Value,
                Subtotal =
                    Enum(dataField, "subtotal", PivotXmlEnums.ParseSubtotal) ?? XLPivotSummary.Sum,
                ShowDataAsFormat =
                    Enum(dataField, "showDataAs", PivotXmlEnums.ParseShowDataAs)
                    ?? XLPivotCalculation.Normal,
                BaseField = OptionalInt(dataField, "baseField") ?? -1,
                BaseItem = OptionalUInt(dataField, "baseItem") ?? 1048832,
                NumberFormatValue = numberFormatId is not null
                    ? styles.NumberFormats[numberFormatId.Value]
                    : null,
            };

            xlPivotTable.DataFields.AddField(xlDataField);
        }

        // Load formats
        foreach (XElement format in Children(pivotTable, "formats", "format"))
        {
            XLDxfValue? dxf = OptionalUInt(format, "dxfId") is { } dxfId
                ? sheet.Workbook.Styles.DifferentialFormats[checked((int)dxfId)]
                : null;

            XElement pivotArea =
                format.Element(Main + "pivotArea")
                ?? throw PartStructureException.ExpectedElementNotFound();

            xlPivotTable.AddFormat(
                new XLPivotFormat(LoadPivotArea(pivotArea))
                {
                    Action =
                        Enum(format, "action", PivotXmlEnums.ParseFormatAction)
                        ?? XLPivotFormatAction.Formatting,
                    FormatValue = dxf,
                }
            );
        }

        foreach (
            XElement conditionalFormat in Children(
                pivotTable,
                "conditionalFormats",
                "conditionalFormat"
            )
        )
        {
            uint priority = RequiredUInt(conditionalFormat, "priority");
            XLConditionalFormat format = context.GetPivotCf(sheet.Name, checked((int)priority));
            XLPivotConditionalFormat xlConditionalFormat = new(format)
            {
                Scope =
                    Enum(conditionalFormat, "scope", PivotXmlEnums.ParseCfScope)
                    ?? XLPivotCfScope.SelectedCells,
                Type =
                    Enum(conditionalFormat, "type", PivotXmlEnums.ParseCfRuleType)
                    ?? XLPivotCfRuleType.None,
            };

            foreach (XElement pivotArea in Children(conditionalFormat, "pivotAreas", "pivotArea"))
            {
                xlConditionalFormat.AddArea(LoadPivotArea(pivotArea));
            }

            xlPivotTable.AddConditionalFormat(xlConditionalFormat);
        }

        // TODO: chartFormats
        // pivotHierarchies is OLAP and thus for now out of scope.
        LoadPivotTableStyle(pivotTable.Element(Main + "pivotTableStyleInfo"), xlPivotTable);

        // TODO: filters
        // rowHierarchiesUsage is OLAP and thus for now out of scope.
        // colHierarchiesUsage is OLAP and thus for now out of scope.
        LoadExtensionList(pivotTable, xlPivotTable);

        return xlPivotTable;
    }

    private static XLPivotTable LoadPivotTableAttributes(
        XElement pivotTable,
        XLWorksheet sheet,
        XLPivotCache cache
    )
    {
        // DataPosition attribute is skipped, because it basically represents a field on one of axis.
        // Excel requires that dataPosition and field with index -2 must be in list of respective axis
        // at correct place, otherwise it crashes. To make things simple, we set the value when it is
        // encountered on the correct axis (plus there is a check that field is not used on multiple axes
        // that would cause exception).
        return new XLPivotTable(sheet, cache)
        {
            Name = RequiredString(pivotTable, "name"),
            DataOnRows = Bool(pivotTable, "dataOnRows") ?? false,
            DataPosition = null, // 'data' field is set when during axis loading (if present).
            AutoFormatId = OptionalUInt(pivotTable, "autoFormatId"),
            ApplyNumberFormats = Bool(pivotTable, "applyNumberFormats") ?? false,
            ApplyBorderFormats = Bool(pivotTable, "applyBorderFormats") ?? false,
            ApplyFontFormats = Bool(pivotTable, "applyFontFormats") ?? false,
            ApplyPatternFormats = Bool(pivotTable, "applyPatternFormats") ?? false,
            ApplyAlignmentFormats = Bool(pivotTable, "applyAlignmentFormats") ?? false,
            ApplyWidthHeightFormats = Bool(pivotTable, "applyWidthHeightFormats") ?? false,
            DataCaption = RequiredString(pivotTable, "dataCaption"),
            GrandTotalCaption = pivotTable.Attribute("grandTotalCaption")?.Value,
            ErrorValueReplacement = pivotTable.Attribute("errorCaption")?.Value,
            ShowError = Bool(pivotTable, "showError") ?? false,
            MissingCaption = pivotTable.Attribute("missingCaption")?.Value ?? string.Empty,
            ShowMissing = Bool(pivotTable, "showMissing") ?? true,
            PageStyle = pivotTable.Attribute("pageStyle")?.Value,

            // The attribute is pivotTableStyle, not pivotTableStyleName.
            PivotTableStyleName = pivotTable.Attribute("pivotTableStyle")?.Value,
            VacatedStyle = pivotTable.Attribute("vacatedStyle")?.Value,
            Tag = pivotTable.Attribute("tag")?.Value,
            UpdatedVersion = OptionalByte(pivotTable, "updatedVersion") ?? 0,
            MinRefreshableVersion = OptionalByte(pivotTable, "minRefreshableVersion") ?? 0,
            AsteriskTotals = Bool(pivotTable, "asteriskTotals") ?? false,
            DisplayItemLabels = Bool(pivotTable, "showItems") ?? true,
            EditData = Bool(pivotTable, "editData") ?? false,
            DisableFieldList = Bool(pivotTable, "disableFieldList") ?? false,
            ShowCalculatedMembers = Bool(pivotTable, "showCalcMbrs") ?? true,
            VisualTotals = Bool(pivotTable, "visualTotals") ?? true,
            ShowMultipleLabel = Bool(pivotTable, "showMultipleLabel") ?? true,
            ShowDataDropDown = Bool(pivotTable, "showDataDropDown") ?? true,
            ShowExpandCollapseButtons = Bool(pivotTable, "showDrill") ?? true,
            PrintExpandCollapsedButtons = Bool(pivotTable, "printDrill") ?? false,
            ShowPropertiesInTooltips = Bool(pivotTable, "showMemberPropertyTips") ?? true,
            ShowContextualTooltips = Bool(pivotTable, "showDataTips") ?? true,
            EnableEditingMechanism = Bool(pivotTable, "enableWizard") ?? true,
            EnableShowDetails = Bool(pivotTable, "enableDrill") ?? true,
            EnableFieldProperties = Bool(pivotTable, "enableFieldProperties") ?? true,
            PreserveCellFormatting = Bool(pivotTable, "preserveFormatting") ?? true,
            AutofitColumns = Bool(pivotTable, "useAutoFormatting") ?? false,
            FilterFieldsPageWrap = checked((int)(OptionalUInt(pivotTable, "pageWrap") ?? 0)),
            FilterAreaOrder =
                Bool(pivotTable, "pageOverThenDown") ?? false
                    ? XLFilterAreaOrder.OverThenDown
                    : XLFilterAreaOrder.DownThenOver,
            FilteredItemsInSubtotals = Bool(pivotTable, "subtotalHiddenItems") ?? false,
            ShowGrandTotalsRows = Bool(pivotTable, "rowGrandTotals") ?? true,
            ShowGrandTotalsColumns = Bool(pivotTable, "colGrandTotals") ?? true,
            PrintTitles = Bool(pivotTable, "fieldPrintTitles") ?? false,
            RepeatRowLabels = Bool(pivotTable, "itemPrintTitles") ?? false,
            MergeAndCenterWithLabels = Bool(pivotTable, "mergeItem") ?? false,
            ShowDropZones = Bool(pivotTable, "showDropZones") ?? true,
            PivotCacheCreatedVersion = OptionalByte(pivotTable, "createdVersion") ?? 0,
            RowLabelIndent = checked((int)(OptionalUInt(pivotTable, "indent") ?? 1)),
            ShowEmptyItemsOnRows = Bool(pivotTable, "showEmptyRow") ?? false,
            ShowEmptyItemsOnColumns = Bool(pivotTable, "showEmptyCol") ?? false,
            DisplayCaptionsAndDropdowns = Bool(pivotTable, "showHeaders") ?? true,
            Compact = Bool(pivotTable, "compact") ?? true,
            Outline = Bool(pivotTable, "outline") ?? false,
            OutlineData = Bool(pivotTable, "outlineData") ?? false,
            CompactData = Bool(pivotTable, "compactData") ?? true,
            Published = Bool(pivotTable, "published") ?? false,
            ClassicPivotTableLayout = Bool(pivotTable, "gridDropZones") ?? false,
            StopImmersiveUi = Bool(pivotTable, "immersive") ?? true,
            AllowMultipleFilters = Bool(pivotTable, "multipleFieldFilters") ?? true,
            ChartFormat = OptionalUInt(pivotTable, "chartFormat") ?? 0,
            RowHeaderCaption = pivotTable.Attribute("rowHeaderCaption")?.Value,
            ColumnHeaderCaption = pivotTable.Attribute("colHeaderCaption")?.Value,
            SortFieldsAtoZ = Bool(pivotTable, "fieldListSortAscending") ?? false,
            MdxSubQueries = Bool(pivotTable, "mdxSubqueries") ?? false,
            UseCustomListsForSorting = Bool(pivotTable, "customListSort") ?? true,
        };
    }

    private static XLPivotTableField LoadPivotField(
        XElement pivotField,
        XLPivotTable xlPivotTable,
        XLWorkbookStyles styles
    )
    {
        int? numberFormatId = checked((int?)OptionalUInt(pivotField, "numFmtId"));

        XLPivotTableField xlField = new(xlPivotTable)
        {
            Name = pivotField.Attribute("name")?.Value,
            Axis = Enum(pivotField, "axis", PivotXmlEnums.ParseAxis),
            DataField = Bool(pivotField, "dataField") ?? false,
            SubtotalCaption = pivotField.Attribute("subtotalCaption")?.Value ?? string.Empty,
            ShowDropDowns = Bool(pivotField, "showDropDowns") ?? true,
            HiddenLevel = Bool(pivotField, "hiddenLevel") ?? false,
            UniqueMemberProperty = pivotField.Attribute("uniqueMemberProperty")?.Value,
            Compact = Bool(pivotField, "compact") ?? true,
            AllDrilled = Bool(pivotField, "allDrilled") ?? false,
            NumberFormatValue = numberFormatId is not null
                ? styles.NumberFormats[numberFormatId.Value]
                : null,
            Outline = Bool(pivotField, "outline") ?? true,
            SubtotalTop = Bool(pivotField, "subtotalTop") ?? true,
            DragToRow = Bool(pivotField, "dragToRow") ?? true,
            DragToColumn = Bool(pivotField, "dragToCol") ?? true,
            MultipleItemSelectionAllowed =
                Bool(pivotField, "multipleItemSelectionAllowed") ?? false,
            DragToPage = Bool(pivotField, "dragToPage") ?? true,
            DragToData = Bool(pivotField, "dragToData") ?? true,
            DragOff = Bool(pivotField, "dragOff") ?? true,
            ShowAll = Bool(pivotField, "showAll") ?? true,
            InsertBlankRow = Bool(pivotField, "insertBlankRow") ?? false,
            ServerField = Bool(pivotField, "serverField") ?? false,
            InsertPageBreak = Bool(pivotField, "insertPageBreak") ?? false,
            AutoShow = Bool(pivotField, "autoShow") ?? false,
            TopAutoShow = Bool(pivotField, "topAutoShow") ?? true,
            HideNewItems = Bool(pivotField, "hideNewItems") ?? false,
            MeasureFilter = Bool(pivotField, "measureFilter") ?? false,
            IncludeNewItemsInFilter = Bool(pivotField, "includeNewItemsInFilter") ?? false,
            ItemPageCount = OptionalUInt(pivotField, "itemPageCount") ?? 10u,
            SortType =
                Enum(pivotField, "sortType", PivotXmlEnums.ParseFieldSort)
                ?? XLPivotSortType.Default,
            DataSourceSort = Bool(pivotField, "dataSourceSort"),
            NonAutoSortDefault = Bool(pivotField, "nonAutoSortDefault") ?? false,
            RankBy = OptionalUInt(pivotField, "rankBy"),
            Subtotals = ReadSubtotals(pivotField, defaultSubtotalDefault: true),
            ShowPropCell = Bool(pivotField, "showPropCell") ?? false,
            ShowPropTip = Bool(pivotField, "showPropTip") ?? false,
            ShowPropAsCaption = Bool(pivotField, "showPropAsCaption") ?? false,
            DefaultAttributeDrillState = Bool(pivotField, "defaultAttributeDrillState") ?? false,
        };

        foreach (XElement item in Children(pivotField, "items", "item"))
        {
            uint? itemIndex = OptionalUInt(item, "x");
            XLPivotFieldItem xlItem = new(
                xlField,
                itemIndex is null ? null : checked((int)itemIndex.Value)
            )
            {
                // Attributes `sd` and `d` were swapped in spec.
                ApproximatelyHasChildren = Bool(item, "c") ?? false,
                Details = Bool(item, "d") ?? false,
                DrillAcrossAttributes = Bool(item, "e") ?? true,
                CalculatedMember = Bool(item, "f") ?? false,
                Hidden = Bool(item, "h") ?? false,
                Missing = Bool(item, "m") ?? false,
                ItemUserCaption = item.Attribute("n")?.Value,
                ValueIsString = Bool(item, "s") ?? false,
                ShowDetails = Bool(item, "sd") ?? true,
                ItemType = Enum(item, "t", PivotXmlEnums.ParseItemType) ?? XLPivotItemType.Data,
            };

            xlField.AddItem(xlItem);
        }

        // TODO: autoSortScope

        // extLst
        xlField.RepeatItemLabels =
            Bool(Extension(pivotField, X14 + "pivotField"), "fillDownLabels") ?? false;

        return xlField;
    }

    private static void LoadAxisFields(
        XElement? fields,
        XLPivotTableAxis axis,
        XLPivotTable xlPivotTable
    )
    {
        if (fields is null)
        {
            return;
        }

        foreach (XElement field in fields.Elements(Main + "field"))
        {
            // Axis can contain 'data' field.
            int fieldIndex = RequiredInt(field, "x");
            if (
                fieldIndex >= xlPivotTable.PivotFields.Count
                || (fieldIndex < 0 && fieldIndex != ValuesFieldIndex)
            )
            {
                throw PartStructureException.InvalidAttributeValue();
            }

            axis.AddField(fieldIndex);
        }
    }

    private static void LoadAxisItems(XElement? axisItems, XLPivotTableAxis axis)
    {
        if (axisItems is null)
        {
            return;
        }

        // Both row and column use RowItem type for axis item, whose element name is `i`.
        List<int> previous = [];
        foreach (XElement axisItem in axisItems.Elements(Main + "i"))
        {
            XLPivotItemType xlItemType =
                Enum(axisItem, "t", PivotXmlEnums.ParseItemType) ?? XLPivotItemType.Data;

            // This is used by the 'data' field.
            int dataFieldIndex = checked((int)(OptionalUInt(axisItem, "i") ?? 0));
            uint repeatedCount = OptionalUInt(axisItem, "r") ?? 0;

            List<int> fieldIndexes =
            [
                .. axisItem.Elements(Main + "x").Select(x => OptionalInt(x, "v") ?? 0),
            ];

            List<int> allFieldIndexes = [.. previous.Take((int)repeatedCount), .. fieldIndexes];
            axis.AddItem(new XLPivotFieldAxisItem(xlItemType, dataFieldIndex, allFieldIndexes));
            previous = allFieldIndexes;
        }
    }

    private static XLPivotArea LoadPivotArea(XElement pivotArea)
    {
        XLPivotArea xlPivotArea = new()
        {
            Field = OptionalInt(pivotArea, "field"),
            Type =
                Enum(pivotArea, "type", PivotXmlEnums.ParsePivotAreaType) ?? XLPivotAreaType.Normal,
            DataOnly = Bool(pivotArea, "dataOnly") ?? true,
            LabelOnly = Bool(pivotArea, "labelOnly") ?? false,
            GrandRow = Bool(pivotArea, "grandRow") ?? false,
            GrandCol = Bool(pivotArea, "grandCol") ?? false,
            CacheIndex = Bool(pivotArea, "cacheIndex") ?? false,
            Outline = Bool(pivotArea, "outline") ?? true,
            Offset = pivotArea.Attribute("offset")?.Value is { } offsetRefText
                ? Area.Parse(offsetRefText)
                : null,
            CollapsedLevelsAreSubtotals = Bool(pivotArea, "collapsedLevelsAreSubtotals") ?? false,
            Axis = Enum(pivotArea, "axis", PivotXmlEnums.ParseAxis),
            FieldPosition = OptionalUInt(pivotArea, "fieldPosition"),
        };

        // Can contain extensions, in theory at least.
        foreach (XElement reference in Children(pivotArea, "references", "reference"))
        {
            xlPivotArea.AddReference(LoadPivotReference(reference));
        }

        return xlPivotArea;
    }

    private static XLPivotReference LoadPivotReference(XElement reference)
    {
        XLPivotReference xlReference = new()
        {
            Field = OptionalUInt(reference, "field"),
            Selected = Bool(reference, "selected") ?? true,
            ByPosition = Bool(reference, "byPosition") ?? false,
            Relative = Bool(reference, "relative") ?? false,
            Subtotals = ReadSubtotals(reference, defaultSubtotalDefault: false),
        };

        // Add indexes after the reference is initialized, so it can check values by
        // cacheIndex/byPosition. A field item is an `x` element, like a member property index.
        foreach (XElement fieldItem in reference.Elements(Main + "x"))
        {
            xlReference.AddFieldItem(RequiredUInt(fieldItem, "v"));
        }

        return xlReference;
    }

    /// <summary>
    /// The subtotal flags, which are the same set of attributes on a pivot field and on a pivot
    /// area reference, except that a field defaults to having the automatic subtotal on.
    /// </summary>
    private static HashSet<XLSubtotalFunction> ReadSubtotals(
        XElement element,
        bool defaultSubtotalDefault
    )
    {
        HashSet<XLSubtotalFunction> subtotals = [];

        Add("defaultSubtotal", XLSubtotalFunction.Automatic, defaultSubtotalDefault);
        Add("sumSubtotal", XLSubtotalFunction.Sum, false);
        Add("countASubtotal", XLSubtotalFunction.Count, false);
        Add("avgSubtotal", XLSubtotalFunction.Average, false);
        Add("maxSubtotal", XLSubtotalFunction.Maximum, false);
        Add("minSubtotal", XLSubtotalFunction.Minimum, false);
        Add("productSubtotal", XLSubtotalFunction.Product, false);
        Add("countSubtotal", XLSubtotalFunction.CountNumbers, false);
        Add("stdDevSubtotal", XLSubtotalFunction.StandardDeviation, false);
        Add("stdDevPSubtotal", XLSubtotalFunction.PopulationStandardDeviation, false);
        Add("varSubtotal", XLSubtotalFunction.Variance, false);
        Add("varPSubtotal", XLSubtotalFunction.PopulationVariance, false);

        return subtotals;

        void Add(string attribute, XLSubtotalFunction function, bool defaultValue)
        {
            if (Bool(element, attribute) ?? defaultValue)
            {
                subtotals.Add(function);
            }
        }
    }

    private static void LoadPivotTableStyle(XElement? pivotTableStyle, XLPivotTable xlPivotTable)
    {
        if (pivotTableStyle is null)
        {
            return;
        }

        xlPivotTable.Theme =
            pivotTableStyle.Attribute("name")?.Value is { } themeName
            && System.Enum.TryParse(themeName, out XLPivotTableTheme xlPivotTableTheme)
                ? xlPivotTableTheme
                : XLPivotTableTheme.None;

        xlPivotTable.ShowRowHeaders = Bool(pivotTableStyle, "showRowHeaders") ?? false;
        xlPivotTable.ShowColumnHeaders = Bool(pivotTableStyle, "showColHeaders") ?? false;
        xlPivotTable.ShowRowStripes = Bool(pivotTableStyle, "showRowStripes") ?? false;
        xlPivotTable.ShowColumnStripes = Bool(pivotTableStyle, "showColStripes") ?? false;

        // Reading showColStripes into ShowLastColumn is what the previous reader did. It looks
        // like a copy and paste slip, but changing it here would change what loads, so it stays
        // until someone decides that separately.
        xlPivotTable.ShowLastColumn = Bool(pivotTableStyle, "showColStripes") ?? false;
    }

    private static void LoadExtensionList(XElement pivotTable, XLPivotTable xlPivotTable)
    {
        XElement? ptExt2010 = Extension(pivotTable, X14 + "pivotTableDefinition");
        if (ptExt2010 is not null)
        {
            xlPivotTable.EnableCellEditing = Bool(ptExt2010, "enableEdit") ?? false;
            xlPivotTable.ShowValuesRow = !(Bool(ptExt2010, "hideValuesRow") ?? false);
        }
    }

    /// <summary>The first extension of the given name in the element's extension list.</summary>
    private static XElement? Extension(XElement element, XName extensionName) =>
        element
            .Element(Main + "extLst")
            ?.Elements(Main + "ext")
            .Select(ext => ext.Element(extensionName))
            .FirstOrDefault(e => e is not null);

    /// <summary>The children of a container element, or nothing when the container is absent.</summary>
    private static IEnumerable<XElement> Children(
        XElement parent,
        string containerName,
        string childName
    ) => parent.Element(Main + containerName)?.Elements(Main + childName) ?? [];

    private static string RequiredString(XElement element, string name) =>
        element.Attribute(name)?.Value ?? throw PartStructureException.MissingAttribute(name);

    private static uint RequiredUInt(XElement element, string name) =>
        OptionalUInt(element, name) ?? throw PartStructureException.MissingAttribute(name);

    private static int RequiredInt(XElement element, string name) =>
        OptionalInt(element, name) ?? throw PartStructureException.MissingAttribute(name);

    /// <summary>
    /// OOXML booleans are written as 1/0 or true/false, and both have to be accepted.
    /// </summary>
    private static bool? Bool(XElement? element, string name) =>
        element?.Attribute(name)?.Value switch
        {
            null => null,
            "1" or "true" or "on" or "True" => true,
            "0" or "false" or "off" or "False" => false,
            _ => throw PartStructureException.InvalidAttributeFormat(),
        };

    private static uint? OptionalUInt(XElement element, string name) =>
        element.Attribute(name)?.Value is not { } value ? null
        : uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint parsed)
            ? parsed
        : throw PartStructureException.InvalidAttributeFormat();

    private static int? OptionalInt(XElement element, string name) =>
        element.Attribute(name)?.Value is not { } value ? null
        : int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? parsed
        : throw PartStructureException.InvalidAttributeFormat();

    private static byte? OptionalByte(XElement element, string name) =>
        element.Attribute(name)?.Value is not { } value ? null
        : byte.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out byte parsed)
            ? parsed
        : throw PartStructureException.InvalidAttributeFormat();

    private static TEnum? Enum<TEnum>(XElement element, string name, Func<string, TEnum> parse)
        where TEnum : struct => element.Attribute(name)?.Value is { } value ? parse(value) : null;
}

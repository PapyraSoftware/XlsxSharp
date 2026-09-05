#nullable disable
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

internal class PivotTableDefinitionPartReader
{
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
        PivotTableCacheDefinitionPart cache = pivotTablePart.PivotTableCacheDefinitionPart;
        string cacheDefinitionRelId = workbookPart.GetIdOfPart(cache);

        XLPivotCache pivotSource = workbook.PivotCachesInternal.FirstOrDefault<XLPivotCache>(ps =>
            ps.WorkbookCacheRelId == cacheDefinitionRelId
        );

        if (pivotSource == null)
        {
            // If it's missing, find a 'similar' pivot cache, i.e. one that's based on the same source range/table
            // Reading the part again is cheap next to the alternative, which is keeping the
            // cache definition's DOM alive just for this fallback.
            IXLPivotSource cacheSource = PivotTableCacheDefinitionPartReader.ReadSource(cache);
            pivotSource = workbook.PivotCachesInternal.FirstOrDefault<XLPivotCache>(ps =>
                ps.Source.Equals(cacheSource)
            );
        }

        PivotTableDefinition pivotTableDefinition = pivotTablePart.PivotTableDefinition;

        XLCell target = ws.FirstCell();
        if (pivotTableDefinition?.Location?.Reference?.HasValue ?? false)
        {
            ws.Range(pivotTableDefinition.Location.Reference.Value).Clear();
            target = ws.Range(pivotTableDefinition.Location.Reference.Value).FirstCell();
        }

        if (target != null && pivotSource != null)
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

#nullable enable
    private static XLPivotTable LoadPivotTableDefinition(
        PivotTableDefinition pivotTable,
        XLWorksheet sheet,
        XLPivotCache cache,
        LoadContext context
    )
    {
        XLWorkbookStyles styles = sheet.Workbook.Styles;

        // Load base attributes
        XLPivotTable xlPivotTable = LoadPivotTableAttributes(pivotTable, sheet, cache);

        // Load location
        Location? location = pivotTable.Location;
        if (location is null)
        {
            throw PartStructureException.ExpectedElementNotFound();
        }

        string referenceText =
            location.Reference?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.Area = Area.Parse(referenceText);
        xlPivotTable.FirstHeaderRow =
            location.FirstHeaderRow?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataRow =
            location.FirstDataRow?.Value ?? throw PartStructureException.MissingAttribute();
        xlPivotTable.FirstDataCol =
            location.FirstDataColumn?.Value ?? throw PartStructureException.MissingAttribute();

        // Skip `rowPageCount` and `colPageCount`, because they are derived from filterAreaOrder, filterFieldsPageWrap and pageField count

        // Load pivot fields
        PivotFields? pivotFields = pivotTable.PivotFields;
        if (pivotFields is not null)
        {
            foreach (PivotField pivotField in pivotFields.Cast<PivotField>())
            {
                xlPivotTable.AddField(LoadPivotField(pivotField, xlPivotTable, styles));
            }
        }

        // Load row axis fields and items
        LoadAxisFields(pivotTable.RowFields, xlPivotTable.RowAxis, xlPivotTable);
        LoadAxisItems(pivotTable.RowItems, xlPivotTable.RowAxis);

        // Load column axis fields and items
        LoadAxisFields(pivotTable.ColumnFields, xlPivotTable.ColumnAxis, xlPivotTable);
        LoadAxisItems(pivotTable.ColumnItems, xlPivotTable.ColumnAxis);

        // Load page fields, i.e. the filters region.
        PageFields? pageFields = pivotTable.PageFields;
        if (pageFields is not null)
        {
            foreach (PageField pageField in pageFields.Cast<PageField>())
            {
                int field =
                    pageField.Field?.Value ?? throw PartStructureException.MissingAttribute();
                int? itemIndex = checked((int?)pageField.Item?.Value);
                int? hierarchyIndex = pageField.Hierarchy?.Value;
                StringValue? hierarchyUniqueName = pageField.Name;
                StringValue? hierarchyDisplayName = pageField.Caption;
                XLPivotPageField xlPageField = new(field)
                {
                    ItemIndex = itemIndex,
                    HierarchyIndex = hierarchyIndex,
                    HierarchyUniqueName = hierarchyUniqueName,
                    HierarchyDisplayName = hierarchyDisplayName,
                };
                xlPivotTable.Filters.AddField(xlPageField);
            }
        }

        // Load data fields.
        DataFields? dataFields = pivotTable.DataFields;
        if (dataFields is not null)
        {
            foreach (DataField dataField in dataFields.Cast<DataField>())
            {
                string? name = dataField.Name?.Value;
                uint field =
                    dataField.Field?.Value ?? throw PartStructureException.MissingAttribute();
                XLPivotSummary subtotal =
                    dataField.Subtotal?.Value.ToXlsxSharp() ?? XLPivotSummary.Sum;
                XLPivotCalculation showDataAsFormat =
                    dataField.ShowDataAs?.Value.ToXlsxSharp() ?? XLPivotCalculation.Normal;
                int baseField = dataField.BaseField?.Value ?? -1;
                uint baseItem = dataField.BaseItem?.Value ?? 1048832;
                int? numberFormatId = checked((int?)dataField.NumberFormatId?.Value);
                XLNumberFormat? numberFormat = numberFormatId is not null
                    ? styles.NumberFormats[numberFormatId.Value]
                    : (XLNumberFormat?)null;
                XLPivotDataField xlDataField = new(xlPivotTable, checked((int)field))
                {
                    DataFieldName = name,
                    Subtotal = subtotal,
                    ShowDataAsFormat = showDataAsFormat,
                    BaseField = baseField,
                    BaseItem = baseItem,
                    NumberFormatValue = numberFormat,
                };
                xlPivotTable.DataFields.AddField(xlDataField);
            }
        }

        // Load formats
        Formats? formats = pivotTable.Formats;
        if (formats is not null)
        {
            foreach (Format format in formats.Cast<Format>())
            {
                XLPivotFormatAction action =
                    format.Action?.Value.ToXlsxSharp() ?? XLPivotFormatAction.Formatting;
                XLDxfValue? dxf = format.FormatId is { } dxfId
                    ? sheet.Workbook.Styles.DifferentialFormats[checked((int)dxfId.Value)]
                    : null;
                PivotArea pivotArea =
                    format.PivotArea ?? throw PartStructureException.ExpectedElementNotFound();
                XLPivotArea xlPivotArea = LoadPivotArea(pivotArea);
                XLPivotFormat xlFormat = new(xlPivotArea) { Action = action, FormatValue = dxf };
                xlPivotTable.AddFormat(xlFormat);
            }
        }

        DocumentFormat.OpenXml.Spreadsheet.ConditionalFormats? conditionalFormats =
            pivotTable.ConditionalFormats;
        if (conditionalFormats is not null)
        {
            foreach (
                ConditionalFormat conditionalFormat in conditionalFormats.Cast<ConditionalFormat>()
            )
            {
                XLPivotCfScope scope =
                    conditionalFormat.Scope?.Value.ToXlsxSharp() ?? XLPivotCfScope.SelectedCells;
                XLPivotCfRuleType type =
                    conditionalFormat.Type?.Value.ToXlsxSharp() ?? XLPivotCfRuleType.None;
                uint priority =
                    conditionalFormat.Priority?.Value
                    ?? throw PartStructureException.MissingAttribute();
                XLConditionalFormat format = context.GetPivotCf(sheet.Name, checked((int)priority));
                XLPivotConditionalFormat xlConditionalFormat = new(format)
                {
                    Scope = scope,
                    Type = type,
                };
                PivotAreas? pivotAreas = conditionalFormat.PivotAreas;
                if (pivotAreas is not null)
                {
                    foreach (PivotArea pivotArea in pivotAreas.Cast<PivotArea>())
                    {
                        XLPivotArea xlPivotArea = LoadPivotArea(pivotArea);
                        xlConditionalFormat.AddArea(xlPivotArea);
                    }
                }

                xlPivotTable.AddConditionalFormat(xlConditionalFormat);
            }
        }

        // TODO: chartFormats
        // pivotHierarchies is OLAP and thus for now out of scope.
        PivotTableStyle? pivotTableStyle = pivotTable.GetFirstChild<PivotTableStyle>();
        LoadPivotTableStyle(pivotTableStyle, xlPivotTable);

        // TODO: filters
        // rowHierarchiesUsage is OLAP and thus for now out of scope.
        // colHierarchiesUsage is OLAP and thus for now out of scope.
        LoadExtensionList(pivotTable, xlPivotTable);

        return xlPivotTable;
    }

    private static XLPivotTable LoadPivotTableAttributes(
        PivotTableDefinition pivotTable,
        XLWorksheet sheet,
        XLPivotCache cache
    )
    {
        string name = pivotTable.Name?.Value ?? throw PartStructureException.MissingAttribute();
        uint cacheId = pivotTable.CacheId?.Value ?? throw PartStructureException.MissingAttribute();
        bool dataOnRows = pivotTable.DataOnRows?.Value ?? false;

        // DataPosition attribute is skipped, because it basically represents a field on one of axis.
        // Excel requires that dataPosition and field with index -2 must be in list of respective axis
        // at correct place, otherwise it crashes. To make things simple, we set the value when it is
        // encountered on the correct axis (plus there is a check that field is not used on multiple axes
        // that would cause exception).
        uint? autoFormatId = pivotTable.AutoFormatId?.Value;
        bool applyNumberFormats = pivotTable.ApplyNumberFormats?.Value ?? false;
        bool applyBorderFormats = pivotTable.ApplyBorderFormats?.Value ?? false;
        bool applyFontFormats = pivotTable.ApplyFontFormats?.Value ?? false;
        bool applyPatternFormats = pivotTable.ApplyPatternFormats?.Value ?? false;
        bool applyAlignmentFormats = pivotTable.ApplyAlignmentFormats?.Value ?? false;
        bool applyWidthHeightFormats = pivotTable.ApplyWidthHeightFormats?.Value ?? false;
        string dataCaption =
            pivotTable.DataCaption?.Value ?? throw PartStructureException.MissingAttribute();
        string? grandTotalCaption = pivotTable.GrandTotalCaption?.Value;
        string? errorCaption = pivotTable.ErrorCaption?.Value;
        bool showError = pivotTable.ShowError?.Value ?? false;
        string missingCaption = pivotTable.MissingCaption?.Value ?? string.Empty;
        bool showMissing = pivotTable.ShowMissing?.Value ?? true;
        string? pageStyle = pivotTable.PageStyle?.Value;
        string? pivotTableStyleName = pivotTable.PivotTableStyleName?.Value;
        string? vacatedStyle = pivotTable.VacatedStyle?.Value;
        string? tag = pivotTable.Tag?.Value;
        byte updatedVersion = pivotTable.UpdatedVersion?.Value ?? 0;
        byte minRefreshableVersion = pivotTable.MinRefreshableVersion?.Value ?? 0;
        bool asteriskTotals = pivotTable.AsteriskTotals?.Value ?? false;
        bool showItems = pivotTable.ShowItems?.Value ?? true;
        bool editData = pivotTable.EditData?.Value ?? false;
        bool disableFieldList = pivotTable.DisableFieldList?.Value ?? false;
        bool showCalculatedMembers = pivotTable.ShowCalculatedMembers?.Value ?? true;
        bool visualTotals = pivotTable.VisualTotals?.Value ?? true;
        bool showMultipleLabel = pivotTable.ShowMultipleLabel?.Value ?? true;
        bool showDataDropDown = pivotTable.ShowDataDropDown?.Value ?? true;
        bool showDrill = pivotTable.ShowDrill?.Value ?? true;
        bool printDrill = pivotTable.PrintDrill?.Value ?? false;
        bool showMemberPropertyTips = pivotTable.ShowMemberPropertyTips?.Value ?? true;
        bool showDataTips = pivotTable.ShowDataTips?.Value ?? true;
        bool enableWizard = pivotTable.EnableWizard?.Value ?? true;
        bool enableDrill = pivotTable.EnableDrill?.Value ?? true;
        bool enableFieldProperties = pivotTable.EnableFieldProperties?.Value ?? true;
        bool preserveFormatting = pivotTable.PreserveFormatting?.Value ?? true;
        bool useAutoFormatting = pivotTable.UseAutoFormatting?.Value ?? false;
        uint pageWrap = pivotTable.PageWrap?.Value ?? 0;
        bool pageOverThenDown = pivotTable.PageOverThenDown?.Value ?? false;
        bool subtotalHiddenItems = pivotTable.SubtotalHiddenItems?.Value ?? false;
        bool rowGrandTotals = pivotTable.RowGrandTotals?.Value ?? true;
        bool columnGrandTotals = pivotTable.ColumnGrandTotals?.Value ?? true;
        bool fieldPrintTitles = pivotTable.FieldPrintTitles?.Value ?? false;
        bool itemPrintTitles = pivotTable.ItemPrintTitles?.Value ?? false;
        bool mergeItem = pivotTable.MergeItem?.Value ?? false;
        bool showDropZones = pivotTable.ShowDropZones?.Value ?? true;
        byte createdVersion = pivotTable.CreatedVersion?.Value ?? 0;
        uint indent = pivotTable.Indent?.Value ?? 1;
        bool showEmptyRow = pivotTable.ShowEmptyRow?.Value ?? false;
        bool showEmptyColumn = pivotTable.ShowEmptyColumn?.Value ?? false;
        bool showHeaders = pivotTable.ShowHeaders?.Value ?? true;
        bool compact = pivotTable.Compact?.Value ?? true;
        bool outline = pivotTable.Outline?.Value ?? false;
        bool outlineData = pivotTable.OutlineData?.Value ?? false;
        bool compactData = pivotTable.CompactData?.Value ?? true;
        bool published = pivotTable.Published?.Value ?? false;
        bool gridDropZones = pivotTable.GridDropZones?.Value ?? false;
        bool stopImmersiveUi = pivotTable.StopImmersiveUi?.Value ?? true;
        bool multipleFieldFilters = pivotTable.MultipleFieldFilters?.Value ?? true;
        uint chartFormat = pivotTable.ChartFormat?.Value ?? 0;
        string? rowHeaderCaption = pivotTable.RowHeaderCaption?.Value;
        string? columnHeaderCaption = pivotTable.ColumnHeaderCaption?.Value;
        bool fieldListSortAscending = pivotTable.FieldListSortAscending?.Value ?? false;
        bool mdxSubQueries = pivotTable.MdxSubqueries?.Value ?? false;
        bool customSortList = pivotTable.CustomListSort?.Value ?? true;

        XLPivotTable xlPivotTable = new(sheet, cache)
        {
            Name = name,
            DataOnRows = dataOnRows,
            DataPosition = null, // 'data' field is set when during axis loading (if present).
            AutoFormatId = autoFormatId,
            ApplyNumberFormats = applyNumberFormats,
            ApplyBorderFormats = applyBorderFormats,
            ApplyFontFormats = applyFontFormats,
            ApplyPatternFormats = applyPatternFormats,
            ApplyAlignmentFormats = applyAlignmentFormats,
            ApplyWidthHeightFormats = applyWidthHeightFormats,
            DataCaption = dataCaption,
            GrandTotalCaption = grandTotalCaption,
            ErrorValueReplacement = errorCaption,
            ShowError = showError,
            MissingCaption = missingCaption,
            ShowMissing = showMissing,
            PageStyle = pageStyle,
            PivotTableStyleName = pivotTableStyleName,
            VacatedStyle = vacatedStyle,
            Tag = tag,
            UpdatedVersion = updatedVersion,
            MinRefreshableVersion = minRefreshableVersion,
            AsteriskTotals = asteriskTotals,
            DisplayItemLabels = showItems,
            EditData = editData,
            DisableFieldList = disableFieldList,
            ShowCalculatedMembers = showCalculatedMembers,
            VisualTotals = visualTotals,
            ShowMultipleLabel = showMultipleLabel,
            ShowDataDropDown = showDataDropDown,
            ShowExpandCollapseButtons = showDrill,
            PrintExpandCollapsedButtons = printDrill,
            ShowPropertiesInTooltips = showMemberPropertyTips,
            ShowContextualTooltips = showDataTips,
            EnableEditingMechanism = enableWizard,
            EnableShowDetails = enableDrill,
            EnableFieldProperties = enableFieldProperties,
            PreserveCellFormatting = preserveFormatting,
            AutofitColumns = useAutoFormatting,
            FilterFieldsPageWrap = checked((int)pageWrap),
            FilterAreaOrder = pageOverThenDown
                ? XLFilterAreaOrder.OverThenDown
                : XLFilterAreaOrder.DownThenOver,
            FilteredItemsInSubtotals = subtotalHiddenItems,
            ShowGrandTotalsRows = rowGrandTotals,
            ShowGrandTotalsColumns = columnGrandTotals,
            PrintTitles = fieldPrintTitles,
            RepeatRowLabels = itemPrintTitles,
            MergeAndCenterWithLabels = mergeItem,
            ShowDropZones = showDropZones,
            PivotCacheCreatedVersion = createdVersion,
            RowLabelIndent = checked((int)indent),
            ShowEmptyItemsOnRows = showEmptyRow,
            ShowEmptyItemsOnColumns = showEmptyColumn,
            DisplayCaptionsAndDropdowns = showHeaders,
            Compact = compact,
            Outline = outline,
            OutlineData = outlineData,
            CompactData = compactData,
            Published = published,
            ClassicPivotTableLayout = gridDropZones,
            StopImmersiveUi = stopImmersiveUi,
            AllowMultipleFilters = multipleFieldFilters,
            ChartFormat = chartFormat,
            RowHeaderCaption = rowHeaderCaption,
            ColumnHeaderCaption = columnHeaderCaption,
            SortFieldsAtoZ = fieldListSortAscending,
            MdxSubQueries = mdxSubQueries,
            UseCustomListsForSorting = customSortList,
        };
        return xlPivotTable;
    }

    private static XLPivotTableField LoadPivotField(
        PivotField pivotField,
        XLPivotTable xlPivotTable,
        XLWorkbookStyles styles
    )
    {
        string? customName = pivotField.Name?.Value;
        XLPivotAxis? axis = pivotField.Axis?.Value.ToXlsxSharp();
        bool dataField = pivotField.DataField?.Value ?? false;
        string? subtotalCaption = pivotField.SubtotalCaption?.Value;
        bool showDropDowns = pivotField.ShowDropDowns?.Value ?? true;
        bool hiddenLevel = pivotField.HiddenLevel?.Value ?? false;
        string? uniqueMemberProperty = pivotField.UniqueMemberProperty?.Value;
        bool compact = pivotField.Compact?.Value ?? true;
        bool allDrilled = pivotField.AllDrilled?.Value ?? false;
        int? numberFormatId = checked((int?)pivotField.NumberFormatId?.Value);
        XLNumberFormat? numberFormat = numberFormatId is not null
            ? styles.NumberFormats[numberFormatId.Value]
            : (XLNumberFormat?)null;
        bool outline = pivotField.Outline?.Value ?? true;
        bool subtotalTop = pivotField.SubtotalTop?.Value ?? true;
        bool dragToRow = pivotField.DragToRow?.Value ?? true;
        bool dragToColumn = pivotField.DragToColumn?.Value ?? true;
        bool multipleItemSelectionAllowed = pivotField.MultipleItemSelectionAllowed?.Value ?? false;
        bool dragToPage = pivotField.DragToPage?.Value ?? true;
        bool dragToData = pivotField.DragToData?.Value ?? true;
        bool dragOff = pivotField.DragOff?.Value ?? true;
        bool showAll = pivotField.ShowAll?.Value ?? true;
        bool insertBlankRow = pivotField.InsertBlankRow?.Value ?? false;
        bool serverField = pivotField.ServerField?.Value ?? false;
        bool insertPageBreak = pivotField.InsertPageBreak?.Value ?? false;
        bool autoShow = pivotField.AutoShow?.Value ?? false;
        bool topAutoShow = pivotField.TopAutoShow?.Value ?? true;
        bool hideNewItems = pivotField.HideNewItems?.Value ?? false;
        bool measureFilter = pivotField.MeasureFilter?.Value ?? false;
        bool includeNewItemsInFilter = pivotField.IncludeNewItemsInFilter?.Value ?? false;
        uint itemPageCount = pivotField.ItemPageCount?.Value ?? 10u;
        XLPivotSortType sortType =
            pivotField.SortType?.Value.ToXlsxSharp() ?? XLPivotSortType.Default;
        bool? dataSourceSort = pivotField.DataSourceSort?.Value;
        bool nonAutoSortDefault = pivotField.NonAutoSortDefault?.Value ?? false;
        uint? rankBy = pivotField.RankBy?.Value;
        bool defaultSubtotal = pivotField.DefaultSubtotal?.Value ?? true;
        bool sumSubtotal = pivotField.SumSubtotal?.Value ?? false;
        bool countASubtotal = pivotField.CountASubtotal?.Value ?? false;
        bool avgSubtotal = pivotField.AverageSubTotal?.Value ?? false;
        bool maxSubtotal = pivotField.MaxSubtotal?.Value ?? false;
        bool minSubtotal = pivotField.MinSubtotal?.Value ?? false;
        bool productSubtotal = pivotField.ApplyProductInSubtotal?.Value ?? false;
        bool countSubtotal = pivotField.CountSubtotal?.Value ?? false;
        bool stdDevSubtotal = pivotField.ApplyStandardDeviationInSubtotal?.Value ?? false;
        bool stdDevPSubtotal = pivotField.ApplyStandardDeviationPInSubtotal?.Value ?? false;
        bool varSubtotal = pivotField.ApplyVarianceInSubtotal?.Value ?? false;
        bool varPSubtotal = pivotField.ApplyVariancePInSubtotal?.Value ?? false;
        bool showPropCell = pivotField.ShowPropCell?.Value ?? false;
        bool showPropTip = pivotField.ShowPropertyTooltip?.Value ?? false;
        bool showPropAsCaption = pivotField.ShowPropAsCaption?.Value ?? false;
        bool defaultAttributeDrillState = pivotField.DefaultAttributeDrillState?.Value ?? false;

        HashSet<XLSubtotalFunction> subtotals = [];
        if (defaultSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Automatic);
        }

        if (sumSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Sum);
        }

        if (countASubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Count);
        }

        if (avgSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Average);
        }

        if (maxSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Maximum);
        }

        if (minSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Minimum);
        }

        if (productSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Product);
        }

        if (countSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.CountNumbers);
        }

        if (stdDevSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.StandardDeviation);
        }

        if (stdDevPSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.PopulationStandardDeviation);
        }

        if (varSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Variance);
        }

        if (varPSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.PopulationVariance);
        }

        XLPivotTableField xlField = new(xlPivotTable)
        {
            Name = customName,
            Axis = axis,
            DataField = dataField,
            SubtotalCaption = subtotalCaption ?? string.Empty,
            ShowDropDowns = showDropDowns,
            HiddenLevel = hiddenLevel,
            UniqueMemberProperty = uniqueMemberProperty,
            Compact = compact,
            AllDrilled = allDrilled,
            NumberFormatValue = numberFormat,
            Outline = outline,
            SubtotalTop = subtotalTop,
            DragToRow = dragToRow,
            DragToColumn = dragToColumn,
            MultipleItemSelectionAllowed = multipleItemSelectionAllowed,
            DragToPage = dragToPage,
            DragToData = dragToData,
            DragOff = dragOff,
            ShowAll = showAll,
            InsertBlankRow = insertBlankRow,
            ServerField = serverField,
            InsertPageBreak = insertPageBreak,
            AutoShow = autoShow,
            TopAutoShow = topAutoShow,
            HideNewItems = hideNewItems,
            MeasureFilter = measureFilter,
            IncludeNewItemsInFilter = includeNewItemsInFilter,
            ItemPageCount = itemPageCount,
            SortType = sortType,
            DataSourceSort = dataSourceSort,
            NonAutoSortDefault = nonAutoSortDefault,
            RankBy = rankBy,
            Subtotals = subtotals,
            ShowPropCell = showPropCell,
            ShowPropTip = showPropTip,
            ShowPropAsCaption = showPropAsCaption,
            DefaultAttributeDrillState = defaultAttributeDrillState,
        };

        Items? items = pivotField.Items;
        if (items is not null)
        {
            foreach (Item item in items.Cast<Item>())
            {
                // Attributes `sd` and `d` were swapped in spec.
                bool approximatelyHasChildren = item.ChildItems?.Value ?? false;
                bool details = item.Expanded?.Value ?? false;
                bool drillAcrossAttributes = item.DrillAcrossAttributes?.Value ?? true;
                bool calculatedMember = item.Calculated?.Value ?? false;
                bool hidden = item.Hidden?.Value ?? false;
                bool missing = item.Missing?.Value ?? false;
                StringValue? itemUserCaption = item.ItemName;
                bool valueIsString = item.HasStringVlue?.Value ?? false;
                bool showDetails = item.HideDetails?.Value ?? true;
                uint? itemIndex = item.Index?.Value;
                XLPivotItemType itemType =
                    item.ItemType?.Value.ToXlsxSharp() ?? XLPivotItemType.Data;
                XLPivotFieldItem xlItem = new(
                    xlField,
                    itemIndex is null ? null : checked((int)itemIndex.Value)
                )
                {
                    ApproximatelyHasChildren = approximatelyHasChildren,
                    Details = details,
                    DrillAcrossAttributes = drillAcrossAttributes,
                    CalculatedMember = calculatedMember,
                    Hidden = hidden,
                    Missing = missing,
                    ItemUserCaption = itemUserCaption,
                    ValueIsString = valueIsString,
                    ShowDetails = showDetails,
                    ItemType = itemType,
                };

                xlField.AddItem(xlItem);
            }
        }

        // TODO: autoSortScope

        // extLst
        PivotFieldExtensionList? pivotFieldExtensionList =
            pivotField.GetFirstChild<PivotFieldExtensionList>();
        PivotFieldExtension? pivotFieldExtension =
            pivotFieldExtensionList?.GetFirstChild<PivotFieldExtension>();
        DocumentFormat.OpenXml.Office2010.Excel.PivotField? field2010 =
            pivotFieldExtension?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotField>();
        xlField.RepeatItemLabels = field2010?.FillDownLabels?.Value ?? false;

        return xlField;
    }

    private static void LoadAxisFields(
        OpenXmlCompositeElement? fields,
        XLPivotTableAxis axis,
        XLPivotTable xlPivotTable
    )
    {
        if (fields is not null)
        {
            foreach (Field field in fields.Cast<Field>())
            {
                // Axis can contain 'data' field.
                int fieldIndex =
                    field.Index?.Value ?? throw PartStructureException.MissingAttribute();
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
    }

    private static void LoadAxisItems(OpenXmlCompositeElement? axisItems, XLPivotTableAxis axis)
    {
        if (axisItems is not null)
        {
            // Both row and column use RowItem type for axis item.
            List<int> previous = [];
            foreach (RowItem axisItem in axisItems.Cast<RowItem>())
            {
                XLPivotItemType xlItemType =
                    axisItem.ItemType?.Value.ToXlsxSharp() ?? XLPivotItemType.Data;
                int dataFieldIndex = checked((int)(axisItem.Index?.Value ?? 0)); // This is used by 'data' field
                uint repeatedCount = axisItem.RepeatedItemCount?.Value ?? 0;
                List<int> fieldIndexes = [];
                foreach (
                    MemberPropertyIndex dataIndex in axisItem.ChildElements.Cast<MemberPropertyIndex>()
                )
                {
                    fieldIndexes.Add(dataIndex.Val?.Value ?? 0);
                }

                List<int> allFieldIndexes = [.. previous.Take((int)repeatedCount), .. fieldIndexes];
                axis.AddItem(new XLPivotFieldAxisItem(xlItemType, dataFieldIndex, allFieldIndexes));
                previous = allFieldIndexes;
            }
        }
    }

    private static XLPivotArea LoadPivotArea(PivotArea pivotArea)
    {
        int? field = pivotArea.Field?.Value;
        XLPivotAreaType type = pivotArea.Type?.Value.ToXlsxSharp() ?? XLPivotAreaType.Normal;
        bool dataOnly = pivotArea.DataOnly?.Value ?? true;
        bool labelOnly = pivotArea.LabelOnly?.Value ?? false;
        bool grandRow = pivotArea.GrandRow?.Value ?? false;
        bool grandCol = pivotArea.GrandColumn?.Value ?? false;
        bool cacheIndex = pivotArea.CacheIndex?.Value ?? false;
        bool outline = pivotArea.Outline?.Value ?? true;
        Area? offset = pivotArea.Offset?.Value is { } offsetRefText
            ? Area.Parse(offsetRefText)
            : (Area?)null;
        bool collapsedLevelsAreSubtotals = pivotArea.CollapsedLevelsAreSubtotals?.Value ?? false;
        XLPivotAxis? axis = pivotArea.Axis?.Value.ToXlsxSharp();
        uint? fieldPosition = pivotArea.FieldPosition?.Value;
        XLPivotArea xlPivotArea = new()
        {
            Field = field,
            Type = type,
            DataOnly = dataOnly,
            LabelOnly = labelOnly,
            GrandRow = grandRow,
            GrandCol = grandCol,
            CacheIndex = cacheIndex,
            Outline = outline,
            Offset = offset,
            CollapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals,
            Axis = axis,
            FieldPosition = fieldPosition,
        };

        // Can contain extensions, in theory at least.
        PivotAreaReferences? references = pivotArea.PivotAreaReferences;
        if (references is not null)
        {
            foreach (PivotAreaReference reference in references.Cast<PivotAreaReference>())
            {
                xlPivotArea.AddReference(LoadPivotReference(reference));
            }
        }

        return xlPivotArea;
    }

    private static XLPivotReference LoadPivotReference(PivotAreaReference reference)
    {
        uint? field = reference.Field?.Value;
        bool selected = reference.Selected?.Value ?? true;
        bool byPosition = reference.ByPosition?.Value ?? false;
        bool relative = reference.Relative?.Value ?? false;
        bool defaultSubtotal = reference.DefaultSubtotal?.Value ?? false;
        bool sumSubtotal = reference.SumSubtotal?.Value ?? false;
        bool countASubtotal = reference.CountASubtotal?.Value ?? false;
        bool avgSubtotal = reference.AverageSubtotal?.Value ?? false;
        bool maxSubtotal = reference.MaxSubtotal?.Value ?? false;
        bool minSubtotal = reference.MinSubtotal?.Value ?? false;
        bool productSubtotal = reference.ApplyProductInSubtotal?.Value ?? false;
        bool countSubtotal = reference.CountSubtotal?.Value ?? false;
        bool stdDevSubtotal = reference.ApplyStandardDeviationInSubtotal?.Value ?? false;
        bool stdDevPSubtotal = reference.ApplyStandardDeviationPInSubtotal?.Value ?? false;
        bool varSubtotal = reference.ApplyVarianceInSubtotal?.Value ?? false;
        bool varPSubtotal = reference.ApplyVariancePInSubtotal?.Value ?? false;

        HashSet<XLSubtotalFunction> subtotals = [];
        if (defaultSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Automatic);
        }

        if (sumSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Sum);
        }

        if (countASubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Count);
        }

        if (avgSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Average);
        }

        if (maxSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Maximum);
        }

        if (minSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Minimum);
        }

        if (productSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Product);
        }

        if (countSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.CountNumbers);
        }

        if (stdDevSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.StandardDeviation);
        }

        if (stdDevPSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.PopulationStandardDeviation);
        }

        if (varSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.Variance);
        }

        if (varPSubtotal)
        {
            subtotals.Add(XLSubtotalFunction.PopulationVariance);
        }

        XLPivotReference xlReference = new()
        {
            Field = field,
            Selected = selected,
            ByPosition = byPosition,
            Relative = relative,
            Subtotals = subtotals,
        };

        // Add indexes after the reference is initialized, so it can check values by cacheIndex/byPosition.
        foreach (FieldItem fieldItem in reference.OfType<FieldItem>())
        {
            uint fieldItemValue =
                fieldItem.Val?.Value ?? throw PartStructureException.MissingAttribute();
            xlReference.AddFieldItem(fieldItemValue);
        }

        return xlReference;
    }

    private static void LoadPivotTableStyle(
        PivotTableStyle? pivotTableStyle,
        XLPivotTable xlPivotTable
    )
    {
        if (pivotTableStyle is not null)
        {
            xlPivotTable.Theme =
                pivotTableStyle.Name is not null
                && Enum.TryParse<XLPivotTableTheme>(
                    pivotTableStyle.Name,
                    out XLPivotTableTheme xlPivotTableTheme
                )
                    ? xlPivotTableTheme
                    : XLPivotTableTheme.None;
            xlPivotTable.ShowRowHeaders = pivotTableStyle.ShowRowHeaders?.Value ?? false;
            xlPivotTable.ShowColumnHeaders = pivotTableStyle.ShowColumnHeaders?.Value ?? false;
            xlPivotTable.ShowRowStripes = pivotTableStyle.ShowRowStripes?.Value ?? false;
            xlPivotTable.ShowColumnStripes = pivotTableStyle.ShowColumnStripes?.Value ?? false;
            xlPivotTable.ShowLastColumn = pivotTableStyle.ShowColumnStripes?.Value ?? false;
        }
    }

    private static void LoadExtensionList(
        PivotTableDefinition pivotTable,
        XLPivotTable xlPivotTable
    )
    {
        PivotTableDefinitionExtensionList? extList =
            pivotTable.GetFirstChild<PivotTableDefinitionExtensionList>();
        PivotTableDefinitionExtension? ext2010 =
            extList?.GetFirstChild<PivotTableDefinitionExtension>();
        DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition? ptExt2010 =
            ext2010?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Excel.PivotTableDefinition>();
        if (ptExt2010 is not null)
        {
            xlPivotTable.EnableCellEditing = ptExt2010.EnableEdit?.Value ?? false;
            bool hideValuesRow = ptExt2010.HideValuesRow?.Value ?? false;
            xlPivotTable.ShowValuesRow = !hideValuesRow;
        }
    }
}

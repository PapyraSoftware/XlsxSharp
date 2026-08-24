#nullable disable

using System.Diagnostics;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.PivotStyleFormats;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

[DebuggerDisplay("{Name}")]
internal class XLPivotTable : IXLPivotTable
{
    private readonly XLWorksheet _worksheet;
    private string _name;

    /// <summary>
    /// List of all fields in the pivot table, roughly represents <c>pivotTableDefinition.
    /// pivotFields</c>. Contains info about each field, mostly page/axis info (data field can
    /// reference same field multiple times, so it mostly stores data in data fields).
    /// </summary>
    private readonly List<XLPivotTableField> _fields = [];
    private readonly List<XLPivotFormat> _formats = [];
    private readonly List<XLPivotConditionalFormat> _conditionalFormats = [];
    private XLPivotCache _cache;
    private int _filterFieldsPageWrap;
    private bool _outline = true;
    private bool _outlineData = false;
    private bool _compact = true;
    private bool _compactData = true;

    internal XLPivotTable(XLWorksheet worksheet, XLPivotCache cache)
    {
        this._worksheet = worksheet;
        this.Filters = new XLPivotTableFilters(this);
        this.RowAxis = new XLPivotTableAxis(this, XLPivotAxis.AxisRow);
        this.ColumnAxis = new XLPivotTableAxis(this, XLPivotAxis.AxisCol);
        this.DataFields = new XLPivotDataFields(this);
        this.Theme = XLPivotTableTheme.PivotStyleLight16;
        this._cache = cache;

        this.SetExcelDefaults();
    }

    IXLPivotCache IXLPivotTable.PivotCache
    {
        get => this.PivotCache;
        set => this.PivotCache = (XLPivotCache)value;
    }

    public IXLCell TargetCell
    {
        get
        {
            int filterRows = this.Filters.GetSizeWithGap().Height;
            Point tableCorner = this.Area.FirstPoint;
            Point targetPoint = tableCorner.ShiftRow(-filterRows);
            return this._worksheet.Internals.CellsCollection.GetCell(targetPoint);
        }
        set
        {
            int filterRows = this.Filters.GetSizeWithGap().Height;
            Point valuePoint = ((XLCell)value).Point;
            Point tableCorner = valuePoint.ShiftRow(filterRows);
            this.Area = this.Area.At(tableCorner);
        }
    }

    public XLPivotCache PivotCache
    {
        get => this._cache;
        set
        {
            IReadOnlyList<string> oldNames = this._cache.FieldNames;
            this._cache = value;
            this.UpdateCacheFields(oldNames);
        }
    }

    public IXLPivotFields ReportFilters => this.Filters;

    public IXLPivotFields ColumnLabels => this.ColumnAxis;

    public IXLPivotFields RowLabels => this.RowAxis;

    public IXLPivotValues Values => this.DataFields;

    public IEnumerable<IXLPivotField> ImplementedFields
    {
        get
        {
            foreach (IXLPivotField pf in this.ReportFilters)
            {
                yield return pf;
            }

            foreach (IXLPivotField pf in this.RowLabels)
            {
                yield return pf;
            }

            foreach (IXLPivotField pf in this.ColumnLabels)
            {
                yield return pf;
            }
        }
    }

    /// <summary>
    /// Table theme this pivot table will use.
    /// </summary>
    public XLPivotTableTheme Theme { get; set; }

    /// <summary>
    /// All fields reflected in the pivot cache.
    /// Order of fields is same as for in the <see cref="PivotCache"/>.
    /// </summary>
    internal IReadOnlyList<XLPivotTableField> PivotFields => this._fields;

    internal XLPivotTableFilters Filters { get; }

    internal XLPivotTableAxis RowAxis { get; }

    internal XLPivotTableAxis ColumnAxis { get; }

    internal XLPivotDataFields DataFields { get; }

    internal IReadOnlyList<XLPivotFormat> Formats => this._formats;

    internal IReadOnlyList<XLPivotConditionalFormat> ConditionalFormats => this._conditionalFormats;

    public IXLPivotTable CopyTo(IXLCell targetCell)
    {
        XLAddressComparer addressComparer = new(ignoreFixed: true);
        if (addressComparer.Equals(targetCell.Address, this.TargetCell.Address))
        {
            throw new InvalidOperationException("Cannot copy pivot table to the target cell.");
        }

        IXLWorksheet targetSheet = targetCell.Worksheet;

        string pivotTableName = this.Name;

        int i = 0;
        List<string> pivotTableNames = [.. targetSheet.PivotTables.Select(pvt => pvt.Name)];
        while (
            !XlsxSharp.XLHelper.ValidateName(
                "pivot table",
                pivotTableName,
                "",
                pivotTableNames,
                out _
            )
        )
        {
            i++;
            pivotTableName = this.Name + i.ToInvariantString();
        }

        XLPivotTable newPivotTable = (XLPivotTable)
            targetSheet.PivotTables.Add(pivotTableName, targetCell, this.PivotCache);

        newPivotTable.RelId = null;

        static void CopyPivotField(IXLPivotField originalPivotField, IXLPivotField newPivotField)
        {
            newPivotField
                .SetSort(originalPivotField.SortType)
                .SetSubtotalCaption(originalPivotField.SubtotalCaption)
                .SetIncludeNewItemsInFilter(originalPivotField.IncludeNewItemsInFilter)
                .SetRepeatItemLabels(originalPivotField.RepeatItemLabels)
                .SetInsertBlankLines(originalPivotField.InsertBlankLines)
                .SetShowBlankItems(originalPivotField.ShowBlankItems)
                .SetInsertPageBreaks(originalPivotField.InsertPageBreaks)
                .SetCollapsed(originalPivotField.Collapsed);

            if (originalPivotField.SubtotalsAtTop.HasValue)
            {
                newPivotField.SetSubtotalsAtTop(originalPivotField.SubtotalsAtTop.Value);
            }

            newPivotField.AddSelectedValues(originalPivotField.SelectedValues);
        }

        foreach (IXLPivotField rf in this.ReportFilters)
        {
            CopyPivotField(rf, newPivotTable.ReportFilters.Add(rf.SourceName, rf.CustomName));
        }

        foreach (IXLPivotField cl in this.ColumnLabels)
        {
            CopyPivotField(cl, newPivotTable.ColumnLabels.Add(cl.SourceName, cl.CustomName));
        }

        foreach (IXLPivotField rl in this.RowLabels)
        {
            CopyPivotField(rl, newPivotTable.RowLabels.Add(rl.SourceName, rl.CustomName));
        }

        foreach (IXLPivotValue v in this.Values)
        {
            IXLPivotValue pivotValue = newPivotTable
                .Values.Add(v.SourceName, v.CustomName)
                .SetSummaryFormula(v.SummaryFormula)
                .SetCalculation(v.Calculation)
                .SetCalculationItem(v.CalculationItem)
                .SetBaseFieldName(v.BaseFieldName)
                .SetBaseItemValue(v.BaseItemValue);

            pivotValue.NumberFormat.NumberFormatId = v.NumberFormat.NumberFormatId;
            pivotValue.NumberFormat.Format = v.NumberFormat.Format;
        }

        newPivotTable.Title = this.Title;
        newPivotTable.Description = this.Description;
        newPivotTable.ColumnHeaderCaption = this.ColumnHeaderCaption;
        newPivotTable.RowHeaderCaption = this.RowHeaderCaption;
        newPivotTable.MergeAndCenterWithLabels = this.MergeAndCenterWithLabels;
        newPivotTable.RowLabelIndent = this.RowLabelIndent;
        newPivotTable.FilterAreaOrder = this.FilterAreaOrder;
        newPivotTable.FilterFieldsPageWrap = this.FilterFieldsPageWrap;
        newPivotTable.ErrorValueReplacement = this.ErrorValueReplacement;
        newPivotTable.ShowMissing = this.ShowMissing;
        newPivotTable.MissingCaption = this.MissingCaption;
        newPivotTable.AutofitColumns = this.AutofitColumns;
        newPivotTable.PreserveCellFormatting = this.PreserveCellFormatting;
        newPivotTable.ShowGrandTotalsColumns = this.ShowGrandTotalsColumns;
        newPivotTable.ShowGrandTotalsRows = this.ShowGrandTotalsRows;
        newPivotTable.FilteredItemsInSubtotals = this.FilteredItemsInSubtotals;
        newPivotTable.AllowMultipleFilters = this.AllowMultipleFilters;
        newPivotTable.UseCustomListsForSorting = this.UseCustomListsForSorting;
        newPivotTable.ShowExpandCollapseButtons = this.ShowExpandCollapseButtons;
        newPivotTable.ShowContextualTooltips = this.ShowContextualTooltips;
        newPivotTable.ShowPropertiesInTooltips = this.ShowPropertiesInTooltips;
        newPivotTable.DisplayCaptionsAndDropdowns = this.DisplayCaptionsAndDropdowns;
        newPivotTable.ClassicPivotTableLayout = this.ClassicPivotTableLayout;
        newPivotTable.ShowValuesRow = this.ShowValuesRow;
        newPivotTable.ShowEmptyItemsOnColumns = this.ShowEmptyItemsOnColumns;
        newPivotTable.ShowEmptyItemsOnRows = this.ShowEmptyItemsOnRows;
        newPivotTable.DisplayItemLabels = this.DisplayItemLabels;
        newPivotTable.SortFieldsAtoZ = this.SortFieldsAtoZ;
        newPivotTable.PrintExpandCollapsedButtons = this.PrintExpandCollapsedButtons;
        newPivotTable.RepeatRowLabels = this.RepeatRowLabels;
        newPivotTable.PrintTitles = this.PrintTitles;
        newPivotTable.EnableShowDetails = this.EnableShowDetails;
        newPivotTable.EnableCellEditing = this.EnableCellEditing;
        newPivotTable.ShowRowHeaders = this.ShowRowHeaders;
        newPivotTable.ShowColumnHeaders = this.ShowColumnHeaders;
        newPivotTable.ShowRowStripes = this.ShowRowStripes;
        newPivotTable.ShowColumnStripes = this.ShowColumnStripes;
        newPivotTable.Theme = this.Theme;
        // TODO: Copy Styleformats

        return newPivotTable;
    }

    public string Name
    {
        get => this._name;
        set
        {
            if (this._name == value)
            {
                return;
            }

            string oldname = this._name ?? string.Empty;

            if (
                !XlsxSharp.XLHelper.ValidateName(
                    "pivot table",
                    value,
                    oldname,
                    this._worksheet.PivotTables.Select<XLPivotTable, string>(pvt => pvt.Name),
                    out string message
                )
            )
            {
                throw new ArgumentException(message, nameof(value));
            }

            this._name = value;

            if (
                !string.IsNullOrWhiteSpace(oldname)
                && !string.Equals(oldname, this._name, StringComparison.OrdinalIgnoreCase)
            )
            {
                this.Worksheet.PivotTables.Delete(oldname);
                (this.Worksheet.PivotTables as XLPivotTables).Add(this._name, this);
            }
        }
    }

    public IXLPivotTable SetName(string value)
    {
        this.Name = value;
        return this;
    }

    public string Title { get; set; }

    public IXLPivotTable SetTitle(string value)
    {
        this.Title = value;
        return this;
    }

    public string Description { get; set; }

    public IXLPivotTable SetDescription(string value)
    {
        this.Description = value;
        return this;
    }

    public IXLPivotTable SetColumnHeaderCaption(string value)
    {
        this.ColumnHeaderCaption = value;
        return this;
    }

    public IXLPivotTable SetRowHeaderCaption(string value)
    {
        this.RowHeaderCaption = value;
        return this;
    }

    public IXLPivotTable SetMergeAndCenterWithLabels()
    {
        this.MergeAndCenterWithLabels = true;
        return this;
    }

    public IXLPivotTable SetMergeAndCenterWithLabels(bool value)
    {
        this.MergeAndCenterWithLabels = value;
        return this;
    }

    public IXLPivotTable SetRowLabelIndent(int value)
    {
        this.RowLabelIndent = value;
        return this;
    }

    public IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value)
    {
        this.FilterAreaOrder = value;
        return this;
    }

    public IXLPivotTable SetFilterFieldsPageWrap(int value)
    {
        this.FilterFieldsPageWrap = value;
        return this;
    }

    public IXLPivotTable SetErrorValueReplacement(string value)
    {
        this.ErrorValueReplacement = value;
        return this;
    }

    public string EmptyCellReplacement
    {
        get
        {
            if (this.ShowMissing)
            {
                return this.MissingCaption;
            }

            return string.Empty;
        }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this.ShowMissing = false;
                this.MissingCaption = string.Empty;
            }
            else
            {
                this.ShowMissing = true;
                this.MissingCaption = value;
            }
        }
    }

    public IXLPivotTable SetEmptyCellReplacement(string value)
    {
        this.EmptyCellReplacement = value;
        return this;
    }

    public IXLPivotTable SetAutofitColumns()
    {
        this.AutofitColumns = true;
        return this;
    }

    public IXLPivotTable SetAutofitColumns(bool value)
    {
        this.AutofitColumns = value;
        return this;
    }

    public IXLPivotTable SetPreserveCellFormatting()
    {
        this.PreserveCellFormatting = true;
        return this;
    }

    public IXLPivotTable SetPreserveCellFormatting(bool value)
    {
        this.PreserveCellFormatting = value;
        return this;
    }

    public IXLPivotTable SetShowGrandTotalsRows()
    {
        this.ShowGrandTotalsRows = true;
        return this;
    }

    public IXLPivotTable SetShowGrandTotalsRows(bool value)
    {
        this.ShowGrandTotalsRows = value;
        return this;
    }

    public IXLPivotTable SetShowGrandTotalsColumns()
    {
        this.ShowGrandTotalsColumns = true;
        return this;
    }

    public IXLPivotTable SetShowGrandTotalsColumns(bool value)
    {
        this.ShowGrandTotalsColumns = value;
        return this;
    }

    public IXLPivotTable SetFilteredItemsInSubtotals()
    {
        this.FilteredItemsInSubtotals = true;
        return this;
    }

    public IXLPivotTable SetFilteredItemsInSubtotals(bool value)
    {
        this.FilteredItemsInSubtotals = value;
        return this;
    }

    public IXLPivotTable SetAllowMultipleFilters()
    {
        this.AllowMultipleFilters = true;
        return this;
    }

    public IXLPivotTable SetAllowMultipleFilters(bool value)
    {
        this.AllowMultipleFilters = value;
        return this;
    }

    public IXLPivotTable SetUseCustomListsForSorting()
    {
        this.UseCustomListsForSorting = true;
        return this;
    }

    public IXLPivotTable SetUseCustomListsForSorting(bool value)
    {
        this.UseCustomListsForSorting = value;
        return this;
    }

    public IXLPivotTable SetShowExpandCollapseButtons()
    {
        this.ShowExpandCollapseButtons = true;
        return this;
    }

    public IXLPivotTable SetShowExpandCollapseButtons(bool value)
    {
        this.ShowExpandCollapseButtons = value;
        return this;
    }

    public IXLPivotTable SetShowContextualTooltips()
    {
        this.ShowContextualTooltips = true;
        return this;
    }

    public IXLPivotTable SetShowContextualTooltips(bool value)
    {
        this.ShowContextualTooltips = value;
        return this;
    }

    public IXLPivotTable SetShowPropertiesInTooltips()
    {
        this.ShowPropertiesInTooltips = true;
        return this;
    }

    public IXLPivotTable SetShowPropertiesInTooltips(bool value)
    {
        this.ShowPropertiesInTooltips = value;
        return this;
    }

    public IXLPivotTable SetDisplayCaptionsAndDropdowns()
    {
        this.DisplayCaptionsAndDropdowns = true;
        return this;
    }

    public IXLPivotTable SetDisplayCaptionsAndDropdowns(bool value)
    {
        this.DisplayCaptionsAndDropdowns = value;
        return this;
    }

    public IXLPivotTable SetClassicPivotTableLayout()
    {
        this.ClassicPivotTableLayout = true;
        return this;
    }

    public IXLPivotTable SetClassicPivotTableLayout(bool value)
    {
        this.ClassicPivotTableLayout = value;
        return this;
    }

    public bool ShowValuesRow { get; set; }

    public IXLPivotTable SetShowValuesRow()
    {
        this.ShowValuesRow = true;
        return this;
    }

    public IXLPivotTable SetShowValuesRow(bool value)
    {
        this.ShowValuesRow = value;
        return this;
    }

    public IXLPivotTable SetShowEmptyItemsOnRows()
    {
        this.ShowEmptyItemsOnRows = true;
        return this;
    }

    public IXLPivotTable SetShowEmptyItemsOnRows(bool value)
    {
        this.ShowEmptyItemsOnRows = value;
        return this;
    }

    public IXLPivotTable SetShowEmptyItemsOnColumns()
    {
        this.ShowEmptyItemsOnColumns = true;
        return this;
    }

    public IXLPivotTable SetShowEmptyItemsOnColumns(bool value)
    {
        this.ShowEmptyItemsOnColumns = value;
        return this;
    }

    public IXLPivotTable SetDisplayItemLabels()
    {
        this.DisplayItemLabels = true;
        return this;
    }

    public IXLPivotTable SetDisplayItemLabels(bool value)
    {
        this.DisplayItemLabels = value;
        return this;
    }

    public IXLPivotTable SetSortFieldsAtoZ()
    {
        this.SortFieldsAtoZ = true;
        return this;
    }

    public IXLPivotTable SetSortFieldsAtoZ(bool value)
    {
        this.SortFieldsAtoZ = value;
        return this;
    }

    public IXLPivotTable SetPrintExpandCollapsedButtons()
    {
        this.PrintExpandCollapsedButtons = true;
        return this;
    }

    public IXLPivotTable SetPrintExpandCollapsedButtons(bool value)
    {
        this.PrintExpandCollapsedButtons = value;
        return this;
    }

    public IXLPivotTable SetRepeatRowLabels()
    {
        this.RepeatRowLabels = true;
        return this;
    }

    public IXLPivotTable SetRepeatRowLabels(bool value)
    {
        this.RepeatRowLabels = value;
        return this;
    }

    public IXLPivotTable SetPrintTitles()
    {
        this.PrintTitles = true;
        return this;
    }

    public IXLPivotTable SetPrintTitles(bool value)
    {
        this.PrintTitles = value;
        return this;
    }

    public IXLPivotTable SetEnableShowDetails()
    {
        this.EnableShowDetails = true;
        return this;
    }

    public IXLPivotTable SetEnableShowDetails(bool value)
    {
        this.EnableShowDetails = value;
        return this;
    }

    public bool EnableCellEditing { get; set; }

    public IXLPivotTable SetEnableCellEditing()
    {
        this.EnableCellEditing = true;
        return this;
    }

    public IXLPivotTable SetEnableCellEditing(bool value)
    {
        this.EnableCellEditing = value;
        return this;
    }

    public bool ShowRowHeaders { get; set; }

    public IXLPivotTable SetShowRowHeaders()
    {
        this.ShowRowHeaders = true;
        return this;
    }

    public IXLPivotTable SetShowRowHeaders(bool value)
    {
        this.ShowRowHeaders = value;
        return this;
    }

    public bool ShowColumnHeaders { get; set; }

    public IXLPivotTable SetShowColumnHeaders()
    {
        this.ShowColumnHeaders = true;
        return this;
    }

    public IXLPivotTable SetShowColumnHeaders(bool value)
    {
        this.ShowColumnHeaders = value;
        return this;
    }

    public bool ShowRowStripes { get; set; }

    public IXLPivotTable SetShowRowStripes()
    {
        this.ShowRowStripes = true;
        return this;
    }

    public IXLPivotTable SetShowRowStripes(bool value)
    {
        this.ShowRowStripes = value;
        return this;
    }

    public bool ShowColumnStripes { get; set; }

    public IXLPivotTable SetShowColumnStripes()
    {
        this.ShowColumnStripes = true;
        return this;
    }

    public IXLPivotTable SetShowColumnStripes(bool value)
    {
        this.ShowColumnStripes = value;
        return this;
    }

    /// <summary>
    /// Part of the pivot table style.
    /// </summary>
    internal bool ShowLastColumn { get; set; } = false;

    public XLPivotSubtotals Subtotals { get; set; }

    public IXLPivotTable SetSubtotals(XLPivotSubtotals value)
    {
        this.Subtotals = value;
        return this;
    }

    public XLPivotLayout Layout
    {
        set
        {
            switch (value)
            {
                case XLPivotLayout.Compact:
                    this._compact = this._compactData = true;
                    this._outline = this._outlineData = false;
                    break;
                case XLPivotLayout.Outline:
                    this._compact = this._compactData = false;
                    this._outline = this._outlineData = true;
                    break;
                case XLPivotLayout.Tabular:
                    this._compact = this._compactData = false;
                    this._outline = this._outlineData = false;
                    break;
                default:
                    throw new UnreachableException();
            }

            // It is necessary to set layout for each pivot field, even ones that are not displayed on an axis. Without it, the tabular layout
            // doesn't display headers for axis fields and only display one "Column labels" button instead.
            this.PivotFields.ForEach(f => f.SetLayout(value));
        }
    }

    public IXLPivotTable SetLayout(XLPivotLayout value)
    {
        this.Layout = value;
        return this;
    }

    public bool InsertBlankLines
    {
        set => this.ImplementedFields.ForEach(f => f.SetInsertBlankLines(value));
    }

    public IXLPivotTable SetInsertBlankLines()
    {
        this.InsertBlankLines = true;
        return this;
    }

    public IXLPivotTable SetInsertBlankLines(bool value)
    {
        this.InsertBlankLines = value;
        return this;
    }

    internal string RelId { get; set; }
    internal string CacheDefinitionRelId { get; set; }

    private void SetExcelDefaults()
    {
        this.ShowMissing = true;
        this.MissingCaption = string.Empty;
        this.ShowColumnHeaders = true;
        this.ShowRowHeaders = true;

        // source http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
        this.DisplayItemLabels = true; //	Show Item Names
        this.ShowExpandCollapseButtons = true; //	Show Expand Collapse
        this.PrintExpandCollapsedButtons = false; //	Print Drill Indicators
        this.ShowPropertiesInTooltips = true; //	Show Member Property ToolTips
        this.ShowContextualTooltips = true; //	Show ToolTips on Data
        this.EnableShowDetails = true; //	Enable Drill Down
        this.PreserveCellFormatting = true; //	Preserve Formatting
        this.AutofitColumns = false; //	Auto Formatting
        this.FilterAreaOrder = XLFilterAreaOrder.DownThenOver; //	Page Over Then Down
        this.FilteredItemsInSubtotals = false; //	Subtotal Hidden Items
        this.ShowGrandTotalsRows = true; //	Row Grand Totals
        this.ShowGrandTotalsColumns = true; //	Grand Totals On Columns
        this.PrintTitles = false; //	Field Print Titles
        this.RepeatRowLabels = false; //	Item Print Titles
        this.MergeAndCenterWithLabels = false; //	Merge Titles
        this.RowLabelIndent = 1; //	Indentation for Compact Axis
        this.ShowEmptyItemsOnRows = false; //	Show Empty Row
        this.ShowEmptyItemsOnColumns = false; //	Show Empty Column
        this.DisplayCaptionsAndDropdowns = true; //	Show Field Headers
        this.ClassicPivotTableLayout = false; //	Enable Drop Zones
        this.AllowMultipleFilters = true; //	Multiple Field Filters
        this.SortFieldsAtoZ = false; //	Default Sort Order
        this.UseCustomListsForSorting = true; //	Custom List AutoSort
    }

    public IXLWorksheet Worksheet => this._worksheet;

    public IXLPivotTableStyleFormats StyleFormats => new XLPivotTableStyleFormats(this);

    public IEnumerable<IXLPivotStyleFormat> AllStyleFormats
    {
        get
        {
            foreach (IXLPivotStyleFormat styleFormat in this.StyleFormats.RowGrandTotalFormats)
            {
                yield return styleFormat;
            }

            foreach (IXLPivotStyleFormat styleFormat in this.StyleFormats.ColumnGrandTotalFormats)
            {
                yield return styleFormat;
            }

            // TODO: Skipped for now, until I implement stubs
            //foreach (var pivotField in ImplementedFields)
            //{
            //    yield return pivotField.StyleFormats.Subtotal;
            //    yield return pivotField.StyleFormats.Header;
            //    yield return pivotField.StyleFormats.Label;
            //    yield return pivotField.StyleFormats.DataValuesFormat;
            //}
        }
    }

#nullable enable
    internal void AddField(XLPivotTableField field) => this._fields.Add(field);

    internal void AddFormat(XLPivotFormat pivotFormat) => this._formats.Add(pivotFormat);

    internal void AddConditionalFormat(XLPivotConditionalFormat conditionalFormat) =>
        this._conditionalFormats.Add(conditionalFormat);

    #region location

    /// <summary>
    /// Area of a pivot table. Area doesn't include page fields, they are above the area with
    /// one empty row between area and filters.
    /// </summary>
    internal Area Area { get; set; } = new(1, 1, 1, 1);

    /// <summary>
    /// First row of pivot table header, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstHeaderRow { get; set; }

    /// <summary>
    /// First row of pivot table data area, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstDataRow { get; set; }

    /// <summary>
    /// First column of pivot table data area, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstDataCol { get; set; }

    #endregion

    #region Attributes of PivotTableDefinition in same order as XSD

    /// <summary>
    /// Determines the whether 'data' field is on <see cref="RowAxis"/> (<c>true</c>) or
    /// <see cref="ColumnAxis"/>(<c>false</c>).
    /// </summary>
    internal bool DataOnRows { get; set; } = false;

    /// <summary>
    /// Determines the default 'data' field position, when it is automatically added to row/column fields.
    /// 0 = first (e.g. before all column/row fields), 1 = second (i.e. after first row/column field) and so on.
    /// &gt; number of fields or <c>null</c> indicates the last position.
    /// </summary>
    internal int? DataPosition { get; set; }

    /// <summary>
    /// <para>
    /// An identification of legacy table auto-format to apply to the pivot table. The
    /// <c>Apply*Formats</c> properties specifies which parts of auto-format to apply. If
    /// <c>null</c> or <see cref="AutofitColumns"/> is not <c>true</c>, legacy auto-format is
    /// not applied.
    /// </para>
    /// <para>
    /// The value must be less than 21 or greater than 4096 and less than or equal to 4117. See
    /// ISO-29500 Annex G.3 for how auto formats look like.
    /// </para>
    /// </summary>
    internal uint? AutoFormatId { get; init; }

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format number format properties.
    /// </summary>
    internal bool ApplyNumberFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format border properties.
    /// </summary>
    internal bool ApplyBorderFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format font properties.
    /// </summary>
    internal bool ApplyFontFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format pattern properties.
    /// </summary>
    internal bool ApplyPatternFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format alignment properties.
    /// </summary>
    internal bool ApplyAlignmentFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format width/height properties.
    /// </summary>
    internal bool ApplyWidthHeightFormats { get; init; } = false;

    /// <summary>
    /// Initial text of 'data' field. This is doesn't do anything, Excel always displays
    /// dynamically a text 'Values', translated to current culture.
    /// </summary>
    internal string DataCaption { get; set; } = "Values";

    internal string? GrandTotalCaption { get; init; }

    /// <summary>
    /// Text to display when in cells that contain error.
    /// </summary>
    public string? ErrorValueReplacement { get; set; }

    /// <summary>
    /// Flag indicating if <see cref="ErrorValueReplacement"/> should be shown when cell contain an error.
    /// </summary>
    internal bool ShowError { get; init; } = false;

    /// <summary>
    /// Test to display for missing items, when <see cref="ShowMissing"/> is <c>true</c>.
    /// </summary>
    internal string MissingCaption { get; set; }

    /// <summary>
    /// Flag indicating if <see cref="MissingCaption"/> should be shown when cell has no value.
    /// </summary>
    /// <remarks>Doesn't seem to work in Excel.</remarks>
    internal bool ShowMissing { get; set; } = true;

    /// <summary>
    /// Name of style to apply to <see cref="XLPivotPageField"/> items headers in <see cref="XLPivotAxis.AxisPage"/>.
    /// </summary>
    internal string? PageStyle { get; init; }

    /// <remarks>Doesn't seem to work in Excel.</remarks>
    internal string? PivotTableStyleName { get; init; }

    /// <summary>
    /// Name of a style to apply to the cells left blank when a pivot table shrinks during a refresh operation.
    /// </summary>
    internal string? VacatedStyle { get; init; }

    internal string? Tag { get; init; }

    /// <summary>
    /// Version of the application that last updated the pivot table. Application-dependent.
    /// </summary>
    internal byte UpdatedVersion { get; init; }

    /// <summary>
    /// Minimum version of the application required to update the pivot table. Application-dependent.
    /// </summary>
    internal byte MinRefreshableVersion { get; init; }

    /// <remarks>OLAP related.</remarks>
    internal bool AsteriskTotals { get; init; } = false;

    /// <summary>
    /// <para>
    /// Should field items be displayed on the axis despite pivot table not having any value
    /// field? <c>true</c> will display items even without data field, <c>false</c> won't.
    /// </para>
    /// <para>
    /// Example: There is an empty pivot table with no value fields. Add field 'Name'
    /// to row fields. Should names be displayed on row despite not having any value field?
    /// </para>
    /// </summary>
    /// <remarks>Also called ShowItems</remarks>
    public bool DisplayItemLabels { get; set; } = true;

    /// <summary>
    /// Flag indicating if user is allowed to edit cells in data area.
    /// </summary>
    internal bool EditData { get; init; } = false;

    /// <summary>
    /// Flag indicating if UI to modify the fields of pivot table is disabled. In Excel, the
    /// whole field area is hidden.
    /// </summary>
    internal bool DisableFieldList { get; init; } = false;

    /// <remarks>OLAP only.</remarks>
    internal bool ShowCalculatedMembers { get; init; } = true;

    /// <remarks>OLAP only.</remarks>
    internal bool VisualTotals { get; init; } = true;

    /// <summary>
    /// A flag indicating whether a page field that has selected multiple items (but not
    /// necessarily all) display "(multiple items)" instead of "All"? If value is <c>false</c>.
    /// page fields will display "All" regardless of whether only item subset is selected or
    /// all items are selected.
    /// </summary>
    internal bool ShowMultipleLabel { get; init; } = true;

    /// <summary>
    /// Doesn't seem to do anything. Should hide drop down filters.
    /// </summary>
    internal bool ShowDataDropDown { get; init; } = true;

    /// <summary>
    /// A flag indicating whether UI should display collapse/expand (drill) buttons in pivot
    /// table axes.
    /// </summary>
    /// <remarks>Also called ShowDrill.</remarks>
    public bool ShowExpandCollapseButtons { get; set; } = true;

    /// <summary>
    /// A flag indicating whether collapse/expand (drill) buttons in pivot table axes should
    /// be printed.
    /// </summary>
    /// <remarks>Also called PrintDrill.</remarks>
    public bool PrintExpandCollapsedButtons { get; set; } = false;

    /// <remarks>OLAP only. Also called ShowMemberPropertyTips.</remarks>
    public bool ShowPropertiesInTooltips { get; set; }

    /// <summary>
    /// A flag indicating whether UI should display a tooltip on data items of pivot table. The
    /// tooltip contain info about value field name, row/col items used to aggregate the value
    /// ect. Note that this tooltip generally hides cell notes, because mouseover displays data
    /// tool tip, rather than the note.
    /// </summary>
    /// <remarks>Also called ShowDataTips.</remarks>
    public bool ShowContextualTooltips { get; set; }

    /// <summary>
    /// A flag indicating whether UI should provide a mechanism to edit the pivot table. If the
    /// value is <c>false</c>, Excel provides ability to refresh data through context menu, but
    /// ribbon or other options to manipulate field or pivot table settings are not present.
    /// </summary>
    /// <remarks>Also called enableWizard.</remarks>
    internal bool EnableEditingMechanism { get; set; } = true;

    /// <remarks>Likely OLAP only. Do not confuse with collapse/expand buttons.</remarks>
    public bool EnableShowDetails { get; set; } = true;

    /// <summary>
    /// A flag indicating whether the user is prevented from displaying PivotField properties.
    /// Not very consistent in Excel, e.g. can't display field properties through context menu
    /// of a pivot table, but can display properties menu through context menu in editing wizard.
    /// </summary>
    internal bool EnableFieldProperties { get; init; } = true;

    /// <summary>
    /// A flag that indicates whether the formatting applied by the user to the pivot table
    /// cells is preserved on refresh.
    /// </summary>
    /// <remarks>Once again, ISO-29500 is buggy and says the opposite. Also called <em>
    /// PreserveFormatting</em></remarks>
    public bool PreserveCellFormatting { get; set; } = true;

    /// <summary>
    /// A flag that indicates whether legacy auto formatting has been applied to the PivotTable
    /// view.
    /// </summary>
    /// <remarks>Also called UseAutoFormatting.</remarks>
    public bool AutofitColumns { get; set; } = false;

    /// <inheritdoc />
    /// <remarks>Also called PageWrap.</remarks>
    public int FilterFieldsPageWrap
    {
        get => this._filterFieldsPageWrap;
        set
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException();
            }

            this._filterFieldsPageWrap = value;
        }
    }

    /// <inheritdoc />
    /// <remarks>Also called <em>PageOverThenDown</em>.</remarks>
    public XLFilterAreaOrder FilterAreaOrder { get; set; } = XLFilterAreaOrder.DownThenOver;

    /// <summary>
    /// A flag that indicates whether hidden pivot items should be included in subtotal
    /// calculated values. If <c>true</c>, data for hidden items are included in subtotals
    /// calculated values. If <c>false</c>, hidden values are not included in subtotal
    /// calculations.
    /// </summary>
    /// <remarks>Also called <em>SubtotalHiddenItems</em>. OLAP only. Option in Excel is grayed
    ///     out and does nothing. The option is un-grayed out when pivot cache is part of data
    ///     model.</remarks>
    public bool FilteredItemsInSubtotals { get; set; } = false;

    /// <summary>
    /// A flag indicating whether grand totals should be displayed for the PivotTable rows.
    /// </summary>
    /// <remarks>Also called <em>RowGrandTotals</em>.</remarks>
    public bool ShowGrandTotalsRows { get; set; } = true;

    /// <summary>
    /// A flag indicating whether grand totals should be displayed for the PivotTable columns.
    /// </summary>
    /// <remarks>Also called <em>ColumnGrandTotals</em>.</remarks>
    public bool ShowGrandTotalsColumns { get; set; } = true;

    /// <summary>
    /// A flag indicating whether when a field name should be printed on all pages.
    /// </summary>
    /// <remarks>Also called <em>FieldPrintTitles</em>.</remarks>
    public bool PrintTitles { get; set; } = false;

    /// <summary>
    /// A flag indicating whether whether PivotItem names should be repeated at the top of each
    /// printed page (e.g. if axis item spans multiple pages, it will be repeated an all pages).
    /// </summary>
    /// <remarks>Also called <em>ItemPrintTitles</em>.</remarks>
    public bool RepeatRowLabels { get; set; } = false;

    /// <summary>
    /// A flag indicating whether row or column titles that span multiple cells should be
    /// merged into a single cell. Useful only in in tabular layout, titles in other layouts
    /// don't span across multiple cells.
    /// </summary>
    /// <remarks>Also called <em>MergeItem</em>.</remarks>
    public bool MergeAndCenterWithLabels { get; set; } = false;

    /// <summary>
    /// A flag indicating whether UI for the pivot table should display large text in field
    /// drop zones when there are no fields in the data region (e.g. <em>Drop Value Fields
    /// Here</em>). Only works in legacy layout mode (i.e. <see cref="ClassicPivotTableLayout"/>
    /// is <c>true</c>).
    /// </summary>
    internal bool ShowDropZones { get; init; } = true;

    /// <summary>
    /// Specifies the version of the application that created the pivot cache. Application-dependent.
    /// </summary>
    /// <remarks>Also called <em>CreatedVersion</em>.</remarks>
    internal byte PivotCacheCreatedVersion { get; init; } = 0;

    /// <summary>
    /// A row indentation increment for row axis when pivot table is in compact layout. Units
    /// are characters.
    /// </summary>
    /// <remarks>Also called <em>Indent</em>.</remarks>
    public int RowLabelIndent { get; set; } = 1;

    /// <summary>
    /// A flag indicating whether to include empty rows in the pivot table (i.e. row axis items
    /// are blank and data items are blank).
    /// </summary>
    /// <remarks>Also called <em>ShowEmptyRow</em>.</remarks>
    public bool ShowEmptyItemsOnRows { get; set; } = false;

    /// <summary>
    /// A flag indicating whether to include empty columns in the table (i.e. column axis items
    /// are blank and data items are blank).
    /// </summary>
    /// <remarks>Also called <em>ShowEmptyColumn</em>.</remarks>
    public bool ShowEmptyItemsOnColumns { get; set; }

    /// <summary>
    /// A flag indicating whether to show field names on axis. The axis items are still
    /// displayed, only field names are not. The dropdowns next to the axis field names
    /// are also displayed/hidden based on the flag.
    /// </summary>
    /// <remarks>Also called <em>ShowHeaders</em>.</remarks>
    public bool DisplayCaptionsAndDropdowns { get; set; } = true;

    /// <summary>
    /// A flag indicating whether new fields should have their
    /// <see cref="XLPivotTableField.Compact"/> flag set to <c>true</c>. By new, it means field
    /// added to page, axes or data fields, not a new field from cache.
    /// </summary>
    internal bool Compact
    {
        get => this._compact;
        init => this._compact = value;
    }

    /// <summary>
    /// A flag indicating whether new fields should have their
    /// <see cref="XLPivotTableField.Outline"/> flag set to <c>true</c>. By new, it means field
    /// added to page, axes or data fields, not a new field from cache.
    /// </summary>
    internal bool Outline
    {
        get => this._outline;
        init => this._outline = value;
    }

    /// <summary>
    /// <para>
    /// A flag that indicates whether 'data'/-2 fields in the PivotTable should be displayed in
    /// outline next column of the sheet. This is basically an equivalent of
    /// <see cref="XLPivotTableField.Outline"/> property for the 'data' fields, because 'data'
    /// field is implicit.
    /// </para>
    /// <para>
    /// When <c>true</c>, the labels from the next field (as ordered by
    /// <see cref="XLPivotTableAxis.Fields"/> for row or column) are displayed in the next
    /// column. Has no effect if 'data' field is last field.
    /// </para>
    /// </summary>
    /// <remarks>Doesn't seem to do much in column axis, only in row axis. Also, Excel
    ///     sometimes seems to favor <see cref="Outline"/> flag instead (likely some less used
    ///     paths in the Excel code).</remarks>
    internal bool OutlineData
    {
        get => this._outlineData;
        init => this._outlineData = value;
    }

    /// <summary>
    /// <para>
    /// A flag that indicates whether 'data'/-2 fields in the PivotTable should be displayed in
    /// compact mode (=same column of the sheet). This is basically an equivalent of
    /// <see cref="XLPivotTableField.Compact"/> property for the 'data' fields, because 'data'
    /// field is implicit.
    /// </para>
    /// <para>
    /// When <c>true</c>, the labels from the next field (as ordered by
    /// <see cref="XLPivotTableAxis.Fields"/> for row or column) are displayed in the same
    /// column (one row below). Has no effect if 'data' field is last field.
    /// </para>
    /// </summary>
    /// <remarks>Doesn't seem to do much in column axis, only in row axis. Also, Excel
    ///     sometimes seems to favor <see cref="Compact"/> flag instead (likely some less used
    ///     paths in the Excel code).</remarks>
    internal bool CompactData
    {
        get => this._compactData;
        init => this._compactData = value;
    }

    /// <summary>
    /// A flag that indicates whether data fields in the pivot table are published and
    /// available for viewing in a server rendering environment.
    /// </summary>
    /// <remarks>No idea what this does. Likely flag for other components that display table
    ///     on a web page.</remarks>
    internal bool Published { get; init; } = false;

    /// <summary>
    /// A flag that indicates whether to apply the classic layout. Classic layout displays the
    /// grid zones in UI where user can drop fields (unless disabled through
    /// <see cref="ShowDropZones"/>).
    /// </summary>
    /// <remarks>Also called <em>GridDropZones</em>.</remarks>
    public bool ClassicPivotTableLayout { get; set; } = false;

    /// <summary>
    /// Likely a flag whether immersive reader should be turned off. Not sure if immersive
    /// reader was ever used outside Word, though Excel for Web added some support in 2023.
    /// </summary>
    internal bool StopImmersiveUi { get; init; } = true;

    /// <summary>
    /// <para>
    /// A flag indicating whether field can have at most most one filter type used. This flag
    /// doesn't allow multiple filters of same type, only multiple different filter types.
    /// </para>
    /// <para>
    /// If false, field can have at most one filter, if user tries to set multiple, previous
    /// one is cleared.
    /// </para>
    /// </summary>
    /// <remarks>Also called <em>multipleFieldFilters</em>.</remarks>
    public bool AllowMultipleFilters { get; set; } = true;

    /// <summary>
    /// Specifies the next pivot chart formatting identifier to use on the pivot table. First
    /// actually used identifier should be 1. The format is used in <c>/chartSpace/pivotSource/
    /// fmtId/@val</c>.
    /// </summary>
    internal uint ChartFormat { get; init; } = 0;

    /// <summary>
    /// The text that will be displayed in row header in compact mode. It is next to drop down
    /// (if enabled) of a label/values filter for fields (if
    /// <see cref="DisplayCaptionsAndDropdowns"/> is set to <c>true</c>). Use localized text
    /// <em>Row labels</em> if property is not specified.
    /// </summary>
    public string? RowHeaderCaption { get; set; } = null;

    /// <summary>
    /// The text that will be displayed in column header in compact mode. It is next to drop down
    /// (if enabled) of a label/values filter for fields (if
    /// <see cref="DisplayCaptionsAndDropdowns"/> is set to <c>true</c>). Use localized text
    /// <em>Column labels</em> if property is not specified.
    /// </summary>
    public string? ColumnHeaderCaption { get; set; } = null;

    /// <summary>
    /// A flag that controls how are fields sorted in the field list UI. <c>true</c> will
    /// display fields sorted alphabetically, <c>false</c> will display fields in the order
    /// fields appear in <see cref="XLPivotCache"/>. OLAP data sources always use alphabetical
    /// sorting.
    /// </summary>
    /// <remarks>Also called <em>fieldListSortAscending</em>.</remarks>
    public bool SortFieldsAtoZ { get; set; } = false;

    /// <summary>
    /// A flag indicating whether MDX sub-queries are supported by OLAP data provider of this
    /// pivot table.
    /// </summary>
    internal bool MdxSubQueries { get; init; } = false;

    /// <summary>
    /// A flag that indicates whether custom lists are used for sorting items of fields, both
    /// initially when the PivotField is initialized and the PivotItems are ordered by their
    /// captions, and later when the user applies a sort.
    /// </summary>
    /// <remarks>Also called <em>customSortList</em>.</remarks>
    public bool UseCustomListsForSorting { get; set; }

    #endregion

    /// <summary>
    /// Add field to a specific axis (page/row/col). Only modified <see cref="PivotFields"/>, doesn't modify
    /// additional info in <see cref="RowAxis"/>, <see cref="ColumnAxis"/> or <see cref="Filters"/>.
    /// </summary>
    internal FieldIndex AddFieldToAxis(string sourceName, string customName, XLPivotAxis axis)
    {
        // Only slices axes can be added through this method.
        Debug.Assert(axis is XLPivotAxis.AxisCol or XLPivotAxis.AxisRow or XLPivotAxis.AxisPage);
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
        {
            if (axis != XLPivotAxis.AxisRow && axis != XLPivotAxis.AxisCol)
            {
                throw new ArgumentException(
                    "Data field can be used only on row or column axis.",
                    nameof(sourceName)
                );
            }

            if (this.RowAxis.ContainsDataField || this.ColumnAxis.ContainsDataField)
            {
                throw new ArgumentException("Data field is already used.", nameof(sourceName));
            }

            bool isRowAxis = axis == XLPivotAxis.AxisRow;

            this.DataOnRows = isRowAxis;
            this.DataPosition = isRowAxis
                ? this.RowAxis.Fields.Count
                : this.ColumnAxis.Fields.Count;
            this.DataCaption = "Values"; // Custom captions don't do anything.
            return FieldIndex.DataField;
        }

        if (!this._cache.TryGetFieldIndex(sourceName, out int fieldIndex))
        {
            throw new InvalidOperationException($"Field '{sourceName}' not found in pivot cache.");
        }

        // Check actual fields.
        bool customNameUsed = this._fields.Any(f =>
            XlsxSharp.XLHelper.NameComparer.Equals(f.Name, customName)
        );
        if (customNameUsed)
        {
            throw new InvalidOperationException($"Custom name '{customName}' is already used.");
        }

        XLPivotTableField field = this._fields[fieldIndex];
        field.Name = customName;
        field.Axis = axis;

        // If it is an axis, all possible values to field items, because they should be referenced in items.
        // Page field must have default item, otherwise Excel asks for repair.
        XLPivotCacheSharedItems sharedItems = this._cache.GetFieldSharedItems(fieldIndex);
        for (int i = 0; i < sharedItems.Count; ++i)
        {
            field.AddItem(new XLPivotFieldItem(field, i));
        }

        // Subtotal items must be synchronized with subtotals. If field has a an item for
        // subtotal function, but doesn't declare subtotals function, Excel will try to
        // repair workbook. Subtotal items can be in any order.
        foreach (XLSubtotalFunction subtotalFunction in field.Subtotals)
        {
            XLPivotItemType itemType = subtotalFunction switch
            {
                XLSubtotalFunction.Automatic => XLPivotItemType.Default,
                XLSubtotalFunction.Sum => XLPivotItemType.Sum,
                XLSubtotalFunction.Count => XLPivotItemType.CountA,
                XLSubtotalFunction.Average => XLPivotItemType.Avg,
                XLSubtotalFunction.Minimum => XLPivotItemType.Min,
                XLSubtotalFunction.Maximum => XLPivotItemType.Max,
                XLSubtotalFunction.Product => XLPivotItemType.Product,
                XLSubtotalFunction.CountNumbers => XLPivotItemType.Count,
                XLSubtotalFunction.StandardDeviation => XLPivotItemType.StdDev,
                XLSubtotalFunction.PopulationStandardDeviation => XLPivotItemType.StdDevP,
                XLSubtotalFunction.Variance => XLPivotItemType.Var,
                XLSubtotalFunction.PopulationVariance => XLPivotItemType.VarP,
                _ => throw new UnreachableException(),
            };
            field.AddItem(new XLPivotFieldItem(field, null) { ItemType = itemType });
        }

        return fieldIndex;
    }

    internal void RemoveFieldFromAxis(FieldIndex index)
    {
        if (index.IsDataField)
        {
            this.DataOnRows = false;
            this.DataPosition = null;
            this.DataCaption = "Values";
        }
        else
        {
            XLPivotTableField field = this._fields[index];
            field.Name = null;
            field.Axis = null;
            field.DataField = false;
            field.MultipleItemSelectionAllowed = false;
        }
    }

    internal bool TryGetSourceNameFieldIndex(string sourceName, out FieldIndex index)
    {
        if (
            XlsxSharp.XLHelper.NameComparer.Equals(
                sourceName,
                XLConstants.PivotTable.ValuesSentinalLabel
            )
        )
        {
            index = FieldIndex.DataField;
            return true;
        }

        if (this.PivotCache.TryGetFieldIndex(sourceName, out int fldIndex))
        {
            index = fldIndex;
            return true;
        }

        index = default;
        return false;
    }

    internal bool TryGetCustomNameFieldIndex(string customName, out FieldIndex index)
    {
        StringComparer comparer = XlsxSharp.XLHelper.NameComparer;
        if (comparer.Equals(customName, XLConstants.PivotTable.ValuesSentinalLabel))
        {
            index = FieldIndex.DataField;
            return true;
        }

        IReadOnlyList<XLPivotTableField> allFields = this.PivotFields;
        for (int i = 0; i < allFields.Count; ++i)
        {
            if (comparer.Equals(customName, allFields[i].Name))
            {
                index = i;
                return true;
            }
        }

        index = default;
        return false;
    }

    /// <summary>
    /// Refresh cache fields after cache has changed.
    /// </summary>
    internal void UpdateCacheFields(IReadOnlyList<string> oldFieldNames)
    {
        // Should be better, but at least refresh fields. A lot of attributes are not
        // kept/initialized from the table. We can't just reuse original objects, because
        // all indices are wrong. Make a copy and then re-set the original properties that
        // are saved before GC takes them.
        HashSet<string> newNames = new(this.PivotCache.FieldNames, XlsxSharp.XLHelper.NameComparer);

        // Source and custom name might not be valid at this point, so keep them.
        List<(string SourceName, string? CustomName, XLPivotDataField Field)> keptDataFields = [];
        foreach (XLPivotDataField dataField in this.DataFields)
        {
            string oldSourceName = oldFieldNames[dataField.Field];
            if (newNames.Contains(oldSourceName))
            {
                keptDataFields.Add((oldSourceName, dataField.DataFieldName, dataField));
            }
        }

        bool includeValuesField = keptDataFields.Count > 1;
        List<string> keptFilterSourceNames = GetKeptNames(
            this.Filters.Fields.Select(x => (FieldIndex)x.Field).ToList(),
            oldFieldNames,
            newNames,
            includeValuesField
        );
        List<string> keptRowSourceNames = GetKeptNames(
            this.RowAxis.Fields,
            oldFieldNames,
            newNames,
            includeValuesField
        );
        List<string> keptColumnSourceNames = GetKeptNames(
            this.ColumnAxis.Fields,
            oldFieldNames,
            newNames,
            includeValuesField
        );

        this.Filters.Clear();
        this.RowAxis.Clear();
        this.ColumnAxis.Clear();
        this.DataFields.Clear();

        this._fields.Clear();
        foreach (string fieldName in this.PivotCache.FieldNames)
        {
            XLPivotTableField field = new(this) { Compact = this.Compact, Outline = this.Outline };
            this._fields.Add(field);
        }

        foreach (string filterName in keptFilterSourceNames)
        {
            this.Filters.Add(filterName, filterName);
        }

        foreach (string rowName in keptRowSourceNames)
        {
            this.RowAxis.AddField(rowName, rowName);
        }

        foreach (string columnName in keptColumnSourceNames)
        {
            this.ColumnAxis.AddField(columnName, columnName);
        }

        foreach (
            (
                string SourceName,
                string? CustomName,
                XLPivotDataField Field
            ) keptDataField in keptDataFields
        )
        {
            XLPivotDataField dataField = this.DataFields.AddField(
                keptDataField.SourceName,
                keptDataField.CustomName
            );
            dataField.Subtotal = keptDataField.Field.Subtotal;
        }

        static List<string> GetKeptNames(
            IReadOnlyList<FieldIndex> fieldIndexes,
            IReadOnlyList<string> oldNames,
            HashSet<string> newNames,
            bool includeDataField
        )
        {
            List<string> result = [];
            foreach (FieldIndex fieldIndex in fieldIndexes)
            {
                if (fieldIndex.IsDataField && includeDataField)
                {
                    result.Add(XLConstants.PivotTable.ValuesSentinalLabel);
                    continue;
                }

                string oldName = oldNames[fieldIndex];
                if (newNames.Contains(oldName))
                {
                    result.Add(oldName);
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Is field used by any axis (row, column, filter), but not data.
    /// </summary>
    internal bool IsFieldUsedOnAxis(FieldIndex fieldIndex)
    {
        if (fieldIndex.IsDataField)
        {
            return this.DataPosition is not null;
        }

        return this.RowAxis.Fields.Contains(fieldIndex)
            || this.ColumnAxis.Fields.Contains(fieldIndex)
            || this.Filters.Contains(fieldIndex);
    }

    internal int GetFieldIndex(XLPivotTableField field)
    {
        int index = this._fields.IndexOf(field);
        if (index < 0)
        {
            throw new ArgumentException($"Unable to find field '{field.Name}'.");
        }

        return index;
    }
}

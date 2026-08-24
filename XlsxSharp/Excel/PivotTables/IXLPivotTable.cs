#nullable disable

using XlsxSharp.Excel.PivotStyleFormats;
using XlsxSharp.Excel.PivotValues;

namespace XlsxSharp.Excel;

public interface IXLPivotTable
{
    public XLPivotTableTheme Theme { get; set; }

    public IXLPivotFields ReportFilters { get; }

    /// <summary>
    /// Labels displayed in columns (i.e. horizontal axis) of the pivot table.
    /// </summary>
    public IXLPivotFields ColumnLabels { get; }

    /// <summary>
    /// Labels displayed in rows (i.e. vertical axis) of the pivot table.
    /// </summary>
    public IXLPivotFields RowLabels { get; }

    public IXLPivotValues Values { get; }

    public string Name { get; set; }
    public string Title { get; set; }
    public string Description { get; set; }

    public string ColumnHeaderCaption { get; set; }
    public string RowHeaderCaption { get; set; }

    /// <summary>
    /// Top left corner cell of a pivot table. If the pivot table contains filters fields, the target cell is top
    /// left cell of the first filter field.
    /// </summary>
    public IXLCell TargetCell { get; set; }

    /// <summary>
    /// The cache of data for the pivot table. The pivot table is created
    /// from cached data, not up-to-date data in a worksheet.
    /// </summary>
    public IXLPivotCache PivotCache { get; set; }

    public bool MergeAndCenterWithLabels { get; set; } // MergeItem
    public int RowLabelIndent { get; set; } // Indent

    /// <summary>
    /// Filter fields layout setting that indicates layout order of filter fields. The layout
    /// uses <see cref="FilterFieldsPageWrap"/> to determine when to break to a new row or
    /// column. Default value is <see cref="XLFilterAreaOrder.DownThenOver"/>.
    /// </summary>
    public XLFilterAreaOrder FilterAreaOrder { get; set; }

    /// <summary>
    /// Specifies the number of page fields to display before starting another row or column.
    /// Value = 0 means unlimited.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">If value &lt; 0.</exception>
    public int FilterFieldsPageWrap { get; set; } // PageWrap

    public string ErrorValueReplacement { get; set; } // ErrorCaption
    public string EmptyCellReplacement { get; set; } // MissingCaption
    public bool AutofitColumns { get; set; } //UseAutoFormatting
    public bool PreserveCellFormatting { get; set; } // PreserveFormatting

    public bool ShowGrandTotalsRows { get; set; } // RowGrandTotals
    public bool ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
    public bool FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
    public bool AllowMultipleFilters { get; set; } // MultipleFieldFilters
    public bool UseCustomListsForSorting { get; set; } // CustomListSort

    public bool ShowExpandCollapseButtons { get; set; }
    public bool ShowContextualTooltips { get; set; }
    public bool ShowPropertiesInTooltips { get; set; }
    public bool DisplayCaptionsAndDropdowns { get; set; }
    public bool ClassicPivotTableLayout { get; set; }
    public bool ShowValuesRow { get; set; }
    public bool ShowEmptyItemsOnRows { get; set; }
    public bool ShowEmptyItemsOnColumns { get; set; }
    public bool DisplayItemLabels { get; set; }
    public bool SortFieldsAtoZ { get; set; }

    public bool PrintExpandCollapsedButtons { get; set; }
    public bool RepeatRowLabels { get; set; }
    public bool PrintTitles { get; set; }

    public bool EnableShowDetails { get; set; }
    public bool EnableCellEditing { get; set; }

    public IXLPivotTable CopyTo(IXLCell targetCell);

    public IXLPivotTable SetName(string value);

    public IXLPivotTable SetTitle(string value);

    public IXLPivotTable SetDescription(string value);

    public IXLPivotTable SetMergeAndCenterWithLabels();
    public IXLPivotTable SetMergeAndCenterWithLabels(bool value);

    public IXLPivotTable SetRowLabelIndent(int value);

    public IXLPivotTable SetFilterAreaOrder(XLFilterAreaOrder value);

    public IXLPivotTable SetFilterFieldsPageWrap(int value);

    public IXLPivotTable SetErrorValueReplacement(string value);

    public IXLPivotTable SetEmptyCellReplacement(string value);

    public IXLPivotTable SetAutofitColumns();
    public IXLPivotTable SetAutofitColumns(bool value);

    public IXLPivotTable SetPreserveCellFormatting();
    public IXLPivotTable SetPreserveCellFormatting(bool value);

    public IXLPivotTable SetShowGrandTotalsRows();
    public IXLPivotTable SetShowGrandTotalsRows(bool value);

    /// <summary>
    /// Should pivot table display a grand total for each row in the last column of a pivot
    /// table (it will enlarge pivot table for extra column).
    /// </summary>
    /// <remarks>
    /// This API has inverse row/column names than the Excel. Excel: <em>On for rows
    /// </em> should use this method <em>ShowGrandTotalsColumns</em>.
    /// </remarks>
    public IXLPivotTable SetShowGrandTotalsColumns();

    public IXLPivotTable SetShowGrandTotalsColumns(bool value);

    public IXLPivotTable SetFilteredItemsInSubtotals();
    public IXLPivotTable SetFilteredItemsInSubtotals(bool value);

    public IXLPivotTable SetAllowMultipleFilters();
    public IXLPivotTable SetAllowMultipleFilters(bool value);

    public IXLPivotTable SetUseCustomListsForSorting();
    public IXLPivotTable SetUseCustomListsForSorting(bool value);

    public IXLPivotTable SetShowExpandCollapseButtons();
    public IXLPivotTable SetShowExpandCollapseButtons(bool value);

    public IXLPivotTable SetShowContextualTooltips();
    public IXLPivotTable SetShowContextualTooltips(bool value);

    public IXLPivotTable SetShowPropertiesInTooltips();
    public IXLPivotTable SetShowPropertiesInTooltips(bool value);

    public IXLPivotTable SetDisplayCaptionsAndDropdowns();
    public IXLPivotTable SetDisplayCaptionsAndDropdowns(bool value);

    public IXLPivotTable SetClassicPivotTableLayout();
    public IXLPivotTable SetClassicPivotTableLayout(bool value);

    public IXLPivotTable SetShowValuesRow();
    public IXLPivotTable SetShowValuesRow(bool value);

    public IXLPivotTable SetShowEmptyItemsOnRows();
    public IXLPivotTable SetShowEmptyItemsOnRows(bool value);

    public IXLPivotTable SetShowEmptyItemsOnColumns();
    public IXLPivotTable SetShowEmptyItemsOnColumns(bool value);

    public IXLPivotTable SetDisplayItemLabels();
    public IXLPivotTable SetDisplayItemLabels(bool value);

    public IXLPivotTable SetSortFieldsAtoZ();
    public IXLPivotTable SetSortFieldsAtoZ(bool value);

    public IXLPivotTable SetPrintExpandCollapsedButtons();
    public IXLPivotTable SetPrintExpandCollapsedButtons(bool value);

    public IXLPivotTable SetRepeatRowLabels();
    public IXLPivotTable SetRepeatRowLabels(bool value);

    public IXLPivotTable SetPrintTitles();
    public IXLPivotTable SetPrintTitles(bool value);

    public IXLPivotTable SetEnableShowDetails();
    public IXLPivotTable SetEnableShowDetails(bool value);

    public IXLPivotTable SetEnableCellEditing();
    public IXLPivotTable SetEnableCellEditing(bool value);

    public IXLPivotTable SetColumnHeaderCaption(string value);

    public IXLPivotTable SetRowHeaderCaption(string value);

    public bool ShowRowHeaders { get; set; }
    public bool ShowColumnHeaders { get; set; }
    public bool ShowRowStripes { get; set; }
    public bool ShowColumnStripes { get; set; }
    public XLPivotSubtotals Subtotals { get; set; }

    /// <summary>
    /// Set the layout of the pivot table. It also changes layout of all pivot fields.
    /// </summary>
    public XLPivotLayout Layout { set; }

    public bool InsertBlankLines { set; }

    public IXLPivotTable SetShowRowHeaders();
    public IXLPivotTable SetShowRowHeaders(bool value);

    public IXLPivotTable SetShowColumnHeaders();
    public IXLPivotTable SetShowColumnHeaders(bool value);

    public IXLPivotTable SetShowRowStripes();
    public IXLPivotTable SetShowRowStripes(bool value);

    public IXLPivotTable SetShowColumnStripes();
    public IXLPivotTable SetShowColumnStripes(bool value);

    public IXLPivotTable SetSubtotals(XLPivotSubtotals value);

    public IXLPivotTable SetLayout(XLPivotLayout value);

    public IXLPivotTable SetInsertBlankLines();
    public IXLPivotTable SetInsertBlankLines(bool value);

    public IXLWorksheet Worksheet { get; }

    public IXLPivotTableStyleFormats StyleFormats { get; }
}

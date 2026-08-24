#nullable disable

using XlsxSharp.Excel.PivotStyleFormats;

namespace XlsxSharp.Excel;

public enum XLSubtotalFunction
{
    Automatic,
    Sum,
    Count,
    Average,
    Maximum,
    Minimum,
    Product,
    CountNumbers,
    StandardDeviation,
    PopulationStandardDeviation,
    Variance,
    PopulationVariance,
}

public enum XLPivotLayout
{
    Outline,
    Tabular,
    Compact,
}

/// <summary>
/// A fluent API representation of a field on an <see cref="IXLPivotTable.RowLabels">row</see>,
/// <see cref="IXLPivotTable.ColumnLabels">column</see> or <see cref="IXLPivotTable.ReportFilters">
/// filter</see> axis of a <see cref="IXLPivotTable"/>.
/// </summary>
/// <remarks>
/// If the field is a 'data' field, a lot of properties don't make sense and can't be set. In
/// such case, the setter will throw <exception cref="InvalidOperationException"/> and getter
/// will return default value for the field.
/// </remarks>
public interface IXLPivotField
{
    /// <summary>
    /// <para>
    /// Name of the field in a pivot table <see cref="IXLPivotTable.PivotCache"/>. If the field
    /// is 'data' field, return <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/>.
    /// </para>
    /// <para>
    /// Note that field name in pivot cache is generally same as in the source data range, but
    /// not always. Field names are unique in the cache and if the source data range contains
    /// duplicate column names, the cache will rename them to keep all names unique.
    /// </para>
    /// </summary>
    public string SourceName { get; }

    /// <summary>
    /// <see cref="CustomName"/> of the field in the pivot table. Custom name is a unique
    /// across all fields used in the pivot table (e.g. if same field is added to values area
    /// multiple times, it must have custom name, e.g. <c>Sum1 of Field</c>,
    /// <c>Sum2 of Field</c>).
    /// </summary>
    /// <exception cref="ArgumentException">When setting name to a name that is already used by
    ///     another field.</exception>
    public string CustomName { get; set; }

    public string SubtotalCaption { get; set; }

    /// <summary>
    /// Get subtotals of the field. The content of the collection depends on the type of subtotal:
    /// <list type="bullet">
    ///   <item>None - the collection is empty.</item>
    ///   <item>Automatic - the collection contains one element with function <see cref="XLSubtotalFunction.Automatic"/>.</item>
    ///   <item>Custom - the collection contains a set of functions (at least one) except the <see cref="XLSubtotalFunction.Automatic"/>.</item>
    /// </list>
    /// </summary>
    public IReadOnlyCollection<XLSubtotalFunction> Subtotals { get; }

    public bool IncludeNewItemsInFilter { get; set; }
    public bool Outline { get; set; }
    public bool Compact { get; set; }
    public bool? SubtotalsAtTop { get; set; }
    public bool RepeatItemLabels { get; set; }
    public bool InsertBlankLines { get; set; }
    public bool ShowBlankItems { get; set; }
    public bool InsertPageBreaks { get; set; }

    /// <summary>
    /// Are all items of the field collapsed?
    /// </summary>
    /// <remarks>
    /// If only a subset of items is collapsed, getter returns <c>false</c>.
    /// </remarks>
    public bool Collapsed { get; set; }

    public XLPivotSortType SortType { get; set; }

    /// <inheritdoc cref="CustomName"/>
    public IXLPivotField SetCustomName(string value);

    public IXLPivotField SetSubtotalCaption(string value);

    public IXLPivotField AddSubtotal(XLSubtotalFunction value);

    public IXLPivotField SetIncludeNewItemsInFilter(bool value = true);

    public IXLPivotField SetLayout(XLPivotLayout value);

    public IXLPivotField SetSubtotalsAtTop(bool value = true);

    public IXLPivotField SetRepeatItemLabels(bool value = true);

    public IXLPivotField SetInsertBlankLines(bool value = true);

    public IXLPivotField SetShowBlankItems(bool value = true);

    public IXLPivotField SetInsertPageBreaks(bool value = true);

    public IXLPivotField SetCollapsed(bool value = true);

    public IXLPivotField SetSort(XLPivotSortType value);

    /// <summary>
    /// Selected values for <see cref="IXLPivotTable.ReportFilters"/> filter of the pivot
    /// table. Empty for non-filter fields.
    /// </summary>
    public IReadOnlyList<XLCellValue> SelectedValues { get; }

    /// <summary>
    /// Add a value to selected values of a filter field (<see cref="IXLPivotTable.ReportFilters"/>).
    /// Doesn't do anything, if this field is not a filter fields.
    /// </summary>
    public IXLPivotField AddSelectedValue(XLCellValue value);

    /// <summary>
    /// Add a values to a selected values of a filter field. Doesn't do anything if this field
    /// is not a filter fields.
    /// </summary>
    public IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values);

    public IXLPivotFieldStyleFormats StyleFormats { get; }

    public bool IsOnRowAxis { get; }
    public bool IsOnColumnAxis { get; }
    public bool IsInFilterList { get; }

    /// <summary>
    /// Index of a field in <see cref="XLPivotTable.PivotFields">all pivot fields</see> or -2
    /// for data field.
    /// </summary>
    public int Offset { get; }
}

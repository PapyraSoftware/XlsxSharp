#nullable disable

namespace XlsxSharp.Excel.Tables;

public enum XLTotalsRowFunction
{
    None,
    Sum,
    Minimum,
    Maximum,
    Average,
    Count,
    CountNumbers,
    StandardDeviation,
    Variance,
    Custom,
}

public interface IXLTableField
{
    /// <summary>
    /// Gets the corresponding column for this table field.
    /// Includes the header and footer cells
    /// </summary>
    /// <value>
    /// The column.
    /// </value>
    public IXLRangeColumn Column { get; }

    /// <summary>
    /// Gets the collection of data cells for this field
    /// Excludes the header and footer cells
    /// </summary>
    /// <value>
    /// The data cells
    /// </value>
    public IXLCells DataCells { get; }

    /// <summary>
    /// Gets the footer cell for the table field.
    /// </summary>
    /// <value>
    /// The footer cell. <c>null</c>, if the table
    /// doesn't have set <see cref="IXLTable.ShowTotalsRow"/>.
    /// </value>
    public IXLCell TotalsCell { get; }

    /// <summary>
    /// Gets the header cell for the table field.
    /// </summary>
    /// <value>
    /// The header cell.<c>null</c>, if the table
    /// doesn't have set <see cref="IXLTable.ShowHeaderRow"/>.
    /// </value>
    public IXLCell HeaderCell { get; }

    /// <summary>
    /// Gets the index of the column (0-based).
    /// </summary>
    /// <value>
    /// The index.
    /// </value>
    public int Index { get; }

    /// <summary>
    /// Gets or sets the name/header of this table field.
    /// The corresponding header cell's value will change if you set this.
    /// </summary>
    /// <value>
    /// The name.
    /// </value>
    public string Name { get; set; }

    /// <summary>
    /// Gets the underlying table for this table field.
    /// </summary>
    public IXLTable Table { get; }

    /// <summary>
    /// Gets or sets the totals row formula in A1 format.
    /// </summary>
    /// <value>
    /// The totals row formula a1.
    /// </value>
    public string TotalsRowFormulaA1 { get; set; }

    /// <summary>
    /// Gets or sets the totals row formula in R1C1 format.
    /// </summary>
    /// <value>
    /// The totals row formula r1 c1.
    /// </value>
    public string TotalsRowFormulaR1C1 { get; set; }

    /// <summary>
    /// Gets or sets the totals row function.
    /// </summary>
    /// <value>
    /// The totals row function.
    /// </value>
    public XLTotalsRowFunction TotalsRowFunction { get; set; }

    /// <summary>
    /// Gets or sets the totals row label (the leftmost cell in the totals row).
    /// </summary>
    /// <value>
    /// The totals row label.
    /// </value>
    /// <exception>If the totals row is not displayed for the table.</exception>
    public string TotalsRowLabel { get; set; }

    /// <summary>
    /// Deletes this table field from the table.
    /// </summary>
    public void Delete();

    /// <summary>
    /// Determines whether all cells this table field have a consistent data type.
    /// </summary>
    public bool IsConsistentDataType();

    /// <summary>
    /// Determines whether all cells this table field have a consistent formula.
    /// </summary>
    public bool IsConsistentFormula();

    /// <summary>
    /// Determines whether all cells this table field have a consistent style.
    /// </summary>
    public bool IsConsistentStyle();
}

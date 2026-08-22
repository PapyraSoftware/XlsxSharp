#nullable disable

using System;
using System.Globalization;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;

namespace XlsxSharp.Excel;

public enum XLScope
{
    Workbook,
    Worksheet,
}

public interface IXLRangeBase : IXLAddressable
{
    public IXLWorksheet Worksheet { get; }

    /// <summary>
    /// Sets a value to every cell in this range.
    /// <para>
    /// Setter will clear a formula, if the cell contains a formula.
    /// If the value is a text that starts with a single quote, setter will prefix the value with a single quote through
    /// <see cref="IXLStyle.IncludeQuotePrefix"/> in Excel too and the value of cell is set to to non-quoted text.
    /// </para>
    /// </summary>
    public XLCellValue Value { set; }

    /// <summary>
    ///   Sets the cells' formula with A1 references.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <value>The formula with A1 references.</value>
    public string FormulaA1 { set; }

    /// <summary>
    /// Create an array formula for all cells in the range.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <exception cref="InvalidOperationException">When the range overlaps with a table, pivot table, merged cells or partially overlaps another array formula.</exception>
    public string FormulaArrayA1 { set; }

    /// <summary>
    ///   Sets the cells' formula with R1C1 references.
    /// </summary>
    /// <remarks>
    /// Setter trims the formula and if formula starts with an <c>=</c>, it is removed. If the
    /// formula contains unprefixed future function (e.g. <c>CONCAT</c>), it will be correctly
    /// prefixed (e.g. <c>_xlfn.CONCAT</c>).
    /// </remarks>
    /// <value>The formula with R1C1 references.</value>
    public string FormulaR1C1 { set; }

    public IXLStyle Style { get; set; }

    /// <summary>
    ///   Gets or sets a value indicating whether this cell's text should be shared or not.
    /// </summary>
    /// <value>
    ///   If false the cell's text will not be shared and stored as an inline value.
    /// </value>
    public bool ShareString { set; }

    /// <summary>
    ///   Returns the collection of cells.
    /// </summary>
    public IXLCells Cells();

    public IXLCells Cells(bool usedCellsOnly);

    public IXLCells Cells(bool usedCellsOnly, XLCellsUsedOptions options);

    public IXLCells Cells(string cells);

    public IXLCells Cells(Func<IXLCell, bool> predicate);

    /// <summary>
    ///   Returns the collection of cells that have a value. Formats are ignored.
    /// </summary>
    public IXLCells CellsUsed();

    /// <summary>
    /// Returns the collection of cells that have a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLCells CellsUsed(XLCellsUsedOptions options);

    public IXLCells CellsUsed(Func<IXLCell, bool> predicate);

    public IXLCells CellsUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate);

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    public IXLCells Search(
        string searchText,
        CompareOptions compareOptions = CompareOptions.Ordinal,
        bool searchFormulae = false
    );

    /// <summary>
    ///   Returns the first cell of this range.
    /// </summary>
    public IXLCell FirstCell();

    /// <summary>
    ///   Returns the first non-empty cell with a value of this range. Formats are ignored.
    ///   <para>The cell's address is going to be ([First Row with a value], [First Column with a value])</para>
    /// </summary>
    public IXLCell FirstCellUsed();

    /// <summary>
    /// Returns the first non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLCell FirstCellUsed(XLCellsUsedOptions options);

    public IXLCell FirstCellUsed(Func<IXLCell, bool> predicate);

    /// <summary>
    /// Returns the first non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    /// <param name="predicate">The predicate used to choose cells</param>
    public IXLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate);

    /// <summary>
    ///   Returns the last cell of this range.
    /// </summary>
    public IXLCell LastCell();

    /// <summary>
    ///   Returns the last non-empty cell with a value of this range. Formats are ignored.
    ///   <para>The cell's address is going to be ([Last Row with a value], [Last Column with a value])</para>
    /// </summary>
    public IXLCell LastCellUsed();

    /// <summary>
    /// Returns the last non-empty cell with a value of this range.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLCell LastCellUsed(XLCellsUsedOptions options);

    public IXLCell LastCellUsed(Func<IXLCell, bool> predicate);

    public IXLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate);

    /// <summary>
    ///   Determines whether this range contains the specified range (completely).
    ///   <para>For partial matches use the range.Intersects method.</para>
    /// </summary>
    /// <param name = "rangeAddress">The range address.</param>
    /// <returns>
    ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
    /// </returns>
    public bool Contains(string rangeAddress);

    /// <summary>
    ///   Determines whether this range contains the specified range (completely).
    ///   <para>For partial matches use the range.Intersects method.</para>
    /// </summary>
    /// <param name = "range">The range to match.</param>
    /// <returns>
    ///   <c>true</c> if this range contains the specified range; otherwise, <c>false</c>.
    /// </returns>
    public bool Contains(IXLRangeBase range);

    public bool Contains(IXLCell cell);

    /// <summary>
    ///   Determines whether this range intersects the specified range.
    ///   <para>For whole matches use the range.Contains method.</para>
    /// </summary>
    /// <param name = "rangeAddress">The range address.</param>
    /// <returns>
    ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
    /// </returns>
    public bool Intersects(string rangeAddress);

    /// <summary>
    ///   Determines whether this range contains the specified range.
    ///   <para>For whole matches use the range.Contains method.</para>
    /// </summary>
    /// <param name = "range">The range to match.</param>
    /// <returns>
    ///   <c>true</c> if this range intersects the specified range; otherwise, <c>false</c>.
    /// </returns>
    public bool Intersects(IXLRangeBase range);

    /// <summary>
    ///   Unmerges this range.
    /// </summary>
    public IXLRange Unmerge();

    /// <summary>
    /// Merges this range. Only the top-left cell will have a value, other values will be blank.
    /// </summary>
    public IXLRange Merge();

    public IXLRange Merge(bool checkIntersect);

    /// <summary>
    /// Creates/adds this range to workbook scoped <see cref="IXLDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// </summary>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    public IXLRange AddToNamed(string name);

    /// <summary>
    /// Creates/adds this range to <see cref="IXLDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    /// <param name = "scope">The scope for the named range.</param>
    /// </summary>
    public IXLRange AddToNamed(string name, XLScope scope);

    /// <summary>
    /// Creates/adds this range to <see cref="IXLDefinedNames"/>.
    /// <para>If the named range exists, it will add this range to that named range.</para>
    /// <param name = "name">Name of the defined name, without sheet.</param>
    /// <param name = "scope">The scope for the named range.</param>
    /// <param name = "comment">The comments for the named range.</param>
    /// </summary>
    public IXLRange AddToNamed(string name, XLScope scope, string comment);

    /// <summary>
    /// Clears the contents of this range.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.All);

    /// <summary>
    ///   Deletes the cell comments from this range.
    /// </summary>
    public void DeleteComments();

    /// <summary>
    /// Set value to all cells in the range.
    /// </summary>
    public IXLRangeBase SetValue(XLCellValue value);

    /// <summary>
    ///   Converts this object to a range.
    /// </summary>
    public IXLRange AsRange();

    public bool IsMerged();

    public bool IsEmpty();

    public bool IsEmpty(XLCellsUsedOptions options);

    /// <summary>
    /// Determines whether range address spans the entire column.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire column; otherwise, <c>false</c>.
    /// </returns>
    public bool IsEntireColumn();

    /// <summary>
    /// Determines whether range address spans the entire row.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire row; otherwise, <c>false</c>.
    /// </returns>
    public bool IsEntireRow();

    /// <summary>
    /// Determines whether the range address spans the entire worksheet.
    /// </summary>
    /// <returns>
    ///   <c>true</c> if is entire sheet; otherwise, <c>false</c>.
    /// </returns>
    public bool IsEntireSheet();

    public IXLPivotTable CreatePivotTable(IXLCell targetCell, string name);

    //IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn);

    public IXLAutoFilter SetAutoFilter();

    public IXLAutoFilter SetAutoFilter(bool value);

    /// <summary>
    /// Returns a data validation rule assigned to the range, if any, or creates a new instance of data validation rule if no rule exists.
    /// </summary>
    public IXLDataValidation GetDataValidation();

    /// <summary>
    /// Creates a new data validation rule for the range, replacing the existing one.
    /// </summary>
    public IXLDataValidation CreateDataValidation();

    public IXLConditionalFormat AddConditionalFormat();

    public void Select();

    /// <summary>
    /// Grows this the current range by one cell to each side
    /// </summary>
    public IXLRangeBase Grow();

    /// <summary>
    /// Grows this the current range by the specified number of cells to each side.
    /// </summary>
    /// <param name="growCount">The grow count.</param>
    public IXLRangeBase Grow(int growCount);

    /// <summary>
    /// Shrinks this current range by one cell.
    /// </summary>
    public IXLRangeBase Shrink();

    /// <summary>
    /// Shrinks the current range by the specified number of cells from each side.
    /// </summary>
    /// <param name="shrinkCount">The shrink count.</param>
    public IXLRangeBase Shrink(int shrinkCount);

    /// <summary>
    /// Returns the intersection of this range with another range on the same worksheet.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    /// <returns>The range address of the intersection</returns>
    public IXLRangeAddress Intersection(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    );

    /// <summary>
    /// Returns the set of cells surrounding the current range.
    /// </summary>
    /// <param name="predicate">The predicate to apply on the resulting set of cells.</param>
    public IXLCells SurroundingCells(Func<IXLCell, bool> predicate = null);

    /// <summary>
    /// Calculates the union of two ranges on the same worksheet.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    /// <returns>
    /// The union
    /// </returns>
    public IXLCells Union(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    );

    /// <summary>
    /// Returns all cells in the current range that are not in the other range.
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <param name="thisRangePredicate">Predicate applied to this range's cells.</param>
    /// <param name="otherRangePredicate">Predicate applied to the other range's cells.</param>
    public IXLCells Difference(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    );

    /// <summary>
    /// Returns a range so that its offset from the target base range is equal to the offset of the current range to the source base range.
    /// For example, if the current range is D4:E4, the source base range is A1:C3, then the relative range to the target base range B10:D13 is E14:F14
    /// </summary>
    /// <param name="sourceBaseRange">The source base range.</param>
    /// <param name="targetBaseRange">The target base range.</param>
    /// <returns>The relative range</returns>
    public IXLRangeBase Relative(IXLRangeBase sourceBaseRange, IXLRangeBase targetBaseRange);
}

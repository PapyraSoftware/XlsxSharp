namespace XlsxSharp.Excel.Sort;

/// <summary>
/// A comparer of rows in a range. It uses semantic of a sort feature in Excel.
/// </summary>
/// <remarks>
/// The comparer should work separate from data, but it would necessitate to sort over
/// <see cref="XLRangeRow"/>. That would require to not only instantiate a new object for each
/// sorted row, but since <see cref="XLRangeRow"/>, it would also be be tracked in range
/// repository, slowing each subsequent operation. To improve performance, comparer has
/// reference to underlaying data and compares row numbers that can be stores in a single
/// allocated array of indexes.
/// </remarks>
internal class XLRangeRowsSortComparer : IComparer<int>
{
    private readonly List<(int ColumnNumber, XLCellValueSortComparer Comparer)> _columnComparers;
    private readonly ValueSlice _valueSlice;

    internal XLRangeRowsSortComparer(XLWorksheet sheet, Area sortRange, IXLSortElements sortColumns)
    {
        if (!sortColumns.Any())
        {
            throw new ArgumentException("Empty sort specification.");
        }

        if (sortRange.Width < sortColumns.Max(x => x.ElementNumber))
        {
            throw new ArgumentException("Range has fewer columns that sort specification.");
        }

        this._valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        this._columnComparers =
        [
            .. sortColumns.Select(se =>
                (se.ElementNumber + sortRange.LeftColumn - 1, new XLCellValueSortComparer(se))
            ),
        ];
    }

    public int Compare(int rowNumber1, int rowNumber2)
    {
        foreach ((int columnNumber, XLCellValueSortComparer comparer) in this._columnComparers)
        {
            XLCellValue row1 = this._valueSlice.GetCellValue(new Point(rowNumber1, columnNumber));
            XLCellValue row2 = this._valueSlice.GetCellValue(new Point(rowNumber2, columnNumber));
            int comparison = comparer.Compare(row1, row2);
            if (comparison != 0)
            {
                return comparison;
            }
        }

        // Row sort should be stable, because otherwise we could randomly switch cells
        // with different formats on subsequent sorts. BCL doesn't support in-place
        // stable sort (Array/List.Sort) directly, only LINQ does it (thus extra copy).
        // Note that stable sort has worse worst case O(N*log(N)^2).
        //
        // As a workaround for stable sort, if all values look same, use the order of rows.
        return rowNumber1 - rowNumber2;
    }
}

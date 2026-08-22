using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.Sort;

internal class XLRangeColumnsSortComparer : IComparer<int>
{
    private readonly List<(int RowNumber, XLCellValueSortComparer Comparer)> _rowComparers;
    private readonly ValueSlice _valueSlice;

    internal XLRangeColumnsSortComparer(XLWorksheet sheet, Area sortRange, IXLSortElements sortRows)
    {
        if (!sortRows.Any())
        {
            throw new ArgumentException("Empty sort specification.");
        }

        if (sortRange.Width < sortRows.Max(x => x.ElementNumber))
        {
            throw new ArgumentException("Range has fewer columns that sort specification.");
        }

        this._valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        this._rowComparers =
        [
            .. sortRows.Select(se =>
                (se.ElementNumber + sortRange.TopRow - 1, new XLCellValueSortComparer(se))
            ),
        ];
    }

    public int Compare(int colNumber1, int colNumber2)
    {
        foreach ((int rowNumber, XLCellValueSortComparer comparer) in this._rowComparers)
        {
            XLCellValue col1 = this._valueSlice.GetCellValue(new Point(rowNumber, colNumber1));
            XLCellValue col2 = this._valueSlice.GetCellValue(new Point(rowNumber, colNumber2));
            int comparison = comparer.Compare(col1, col2);
            if (comparison != 0)
            {
                return comparison;
            }
        }

        // Workaround for stable sort, see XLRangeRowsSortComparer.
        return colNumber1 - colNumber2;
    }
}

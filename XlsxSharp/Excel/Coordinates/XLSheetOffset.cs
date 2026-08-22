using System;

namespace XlsxSharp.Excel;

/// <summary>
/// An offset of a cell in a sheet.
/// </summary>
/// <param name="RowOfs">The row offset in number of rows from the original point.</param>
/// <param name="ColOfs">The column offset in number of columns from the original point</param>
internal readonly record struct XLSheetOffset(int RowOfs, int ColOfs) : IComparable<XLSheetOffset>
{
    public int CompareTo(XLSheetOffset other)
    {
        int rowComparison = this.RowOfs.CompareTo(other.RowOfs);
        if (rowComparison != 0)
        {
            return rowComparison;
        }

        return this.ColOfs.CompareTo(other.ColOfs);
    }
}

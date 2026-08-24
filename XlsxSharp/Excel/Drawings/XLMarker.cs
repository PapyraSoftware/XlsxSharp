#nullable disable

using System.Diagnostics;

namespace XlsxSharp.Excel.Drawings;

[DebuggerDisplay("{Address} {Offset}")]
internal class XLMarker
{
    // Using a range to store the location so that it gets added to the range repository
    // and hence will be adjusted when there are insertions / deletions
    private readonly IXLRange rangeCell;

    internal XLMarker(IXLCell cell)
        : this(cell.AsRange(), new System.Drawing.Point(0, 0)) { }

    internal XLMarker(IXLCell cell, System.Drawing.Point offset)
        : this(cell.AsRange(), offset) { }

    private XLMarker(IXLRange rangeCell, System.Drawing.Point offset)
    {
        if (rangeCell.RowCount() != 1 || rangeCell.ColumnCount() != 1)
        {
            throw new ArgumentException("Range should contain only one cell.", nameof(rangeCell));
        }

        this.rangeCell = rangeCell;
        this.Offset = offset;
    }

    public IXLCell Cell => this.rangeCell.FirstCell();

    public int ColumnNumber => this.rangeCell.RangeAddress.FirstAddress.ColumnNumber;
    public System.Drawing.Point Offset { get; set; }
    public int RowNumber => this.rangeCell.RangeAddress.FirstAddress.RowNumber;
}

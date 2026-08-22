#nullable disable

namespace XlsxSharp.Excel.Drawings;

internal class XLDrawingPosition : IXLDrawingPosition
{
    public int Column { get; set; }

    public IXLDrawingPosition SetColumn(int column)
    {
        this.Column = column;
        return this;
    }

    public double ColumnOffset { get; set; }

    public IXLDrawingPosition SetColumnOffset(double columnOffset)
    {
        this.ColumnOffset = columnOffset;
        return this;
    }

    public int Row { get; set; }

    public IXLDrawingPosition SetRow(int row)
    {
        this.Row = row;
        return this;
    }

    public double RowOffset { get; set; }

    public IXLDrawingPosition SetRowOffset(double rowOffset)
    {
        this.RowOffset = rowOffset;
        return this;
    }
}

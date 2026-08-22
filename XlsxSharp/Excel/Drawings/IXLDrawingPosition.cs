#nullable disable

namespace XlsxSharp.Excel.Drawings;

public interface IXLDrawingPosition
{
    public int Column { get; set; }
    public IXLDrawingPosition SetColumn(int column);
    public double ColumnOffset { get; set; }
    public IXLDrawingPosition SetColumnOffset(double columnOffset);

    public int Row { get; set; }
    public IXLDrawingPosition SetRow(int row);
    public double RowOffset { get; set; }
    public IXLDrawingPosition SetRowOffset(double rowOffset);
}

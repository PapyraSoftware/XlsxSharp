#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingSize
{
    public bool AutomaticSize { get; set; }
    public double Height { get; set; }
    public double Width { get; set; }

    public IXLDrawingStyle SetAutomaticSize();
    public IXLDrawingStyle SetAutomaticSize(bool value);
    public IXLDrawingStyle SetHeight(double value);
    public IXLDrawingStyle SetWidth(double value);
}

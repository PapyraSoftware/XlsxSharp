#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingStyle
{
    //IXLDrawingFont Font { get; }
    public IXLDrawingAlignment Alignment { get; }
    public IXLDrawingColorsAndLines ColorsAndLines { get; }
    public IXLDrawingSize Size { get; }
    public IXLDrawingProtection Protection { get; }
    public IXLDrawingProperties Properties { get; }
    public IXLDrawingMargins Margins { get; }
    public IXLDrawingWeb Web { get; }
}

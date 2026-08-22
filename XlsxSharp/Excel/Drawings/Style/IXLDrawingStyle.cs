#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingStyle
{
    //IXLDrawingFont Font { get; }
    IXLDrawingAlignment Alignment { get; }
    IXLDrawingColorsAndLines ColorsAndLines { get; }
    IXLDrawingSize Size { get; }
    IXLDrawingProtection Protection { get; }
    IXLDrawingProperties Properties { get; }
    IXLDrawingMargins Margins { get; }
    IXLDrawingWeb Web { get; }
}

#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingProperties
{
    XLDrawingAnchor Positioning { get; set; }
    IXLDrawingStyle SetPositioning(XLDrawingAnchor value);
}

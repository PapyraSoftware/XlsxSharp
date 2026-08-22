#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingProperties
{
    public XLDrawingAnchor Positioning { get; set; }
    public IXLDrawingStyle SetPositioning(XLDrawingAnchor value);
}

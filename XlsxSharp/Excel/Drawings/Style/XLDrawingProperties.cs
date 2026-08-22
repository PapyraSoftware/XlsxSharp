#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingProperties : IXLDrawingProperties
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingProperties(IXLDrawingStyle style) => this._style = style;

    public XLDrawingAnchor Positioning { get; set; }

    public IXLDrawingStyle SetPositioning(XLDrawingAnchor value)
    {
        this.Positioning = value;
        return this._style;
    }
}

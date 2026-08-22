#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingAlignment : IXLDrawingAlignment
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingAlignment(IXLDrawingStyle style)
    {
        this._style = style;
    }

    public XLDrawingHorizontalAlignment Horizontal { get; set; }

    public IXLDrawingStyle SetHorizontal(XLDrawingHorizontalAlignment value)
    {
        this.Horizontal = value;
        return this._style;
    }

    public XLDrawingVerticalAlignment Vertical { get; set; }

    public IXLDrawingStyle SetVertical(XLDrawingVerticalAlignment value)
    {
        this.Vertical = value;
        return this._style;
    }

    public Boolean AutomaticSize { get; set; }

    public IXLDrawingStyle SetAutomaticSize()
    {
        this.AutomaticSize = true;
        return this._style;
    }

    public IXLDrawingStyle SetAutomaticSize(Boolean value)
    {
        this.AutomaticSize = value;
        return this._style;
    }

    public XLDrawingTextDirection Direction { get; set; }

    public IXLDrawingStyle SetDirection(XLDrawingTextDirection value)
    {
        this.Direction = value;
        return this._style;
    }

    public XLDrawingTextOrientation Orientation { get; set; }

    public IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value)
    {
        this.Orientation = value;
        return this._style;
    }
}

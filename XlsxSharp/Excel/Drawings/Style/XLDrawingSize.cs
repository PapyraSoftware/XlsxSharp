#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingSize : IXLDrawingSize
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingSize(IXLDrawingStyle style)
    {
        this._style = style;
    }

    public Boolean AutomaticSize
    {
        get { return this._style.Alignment.AutomaticSize; }
        set { this._style.Alignment.AutomaticSize = value; }
    }

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

    public Double Height { get; set; }

    public IXLDrawingStyle SetHeight(Double value)
    {
        this.Height = value;
        return this._style;
    }

    public Double Width { get; set; }

    public IXLDrawingStyle SetWidth(Double value)
    {
        this.Width = value;
        return this._style;
    }
}

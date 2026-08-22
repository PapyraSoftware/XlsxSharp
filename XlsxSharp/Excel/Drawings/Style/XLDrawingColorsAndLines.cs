#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingColorsAndLines : IXLDrawingColorsAndLines
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingColorsAndLines(IXLDrawingStyle style)
    {
        this._style = style;
    }

    public XLColor FillColor { get; set; }

    public IXLDrawingStyle SetFillColor(XLColor value)
    {
        this.FillColor = value;
        return this._style;
    }

    public Double FillTransparency { get; set; }

    public IXLDrawingStyle SetFillTransparency(Double value)
    {
        this.FillTransparency = value;
        return this._style;
    }

    public XLColor LineColor { get; set; }

    public IXLDrawingStyle SetLineColor(XLColor value)
    {
        this.LineColor = value;
        return this._style;
    }

    public Double LineTransparency { get; set; }

    public IXLDrawingStyle SetLineTransparency(Double value)
    {
        this.LineTransparency = value;
        return this._style;
    }

    public Double LineWeight { get; set; }

    public IXLDrawingStyle SetLineWeight(Double value)
    {
        this.LineWeight = value;
        return this._style;
    }

    public XLDashStyle LineDash { get; set; }

    public IXLDrawingStyle SetLineDash(XLDashStyle value)
    {
        this.LineDash = value;
        return this._style;
    }

    public XLLineStyle LineStyle { get; set; }

    public IXLDrawingStyle SetLineStyle(XLLineStyle value)
    {
        this.LineStyle = value;
        return this._style;
    }
}

#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public enum XLDashStyle
{
    Solid,
    RoundDot,
    SquareDot,
    Dash,
    DashDot,
    LongDash,
    LongDashDot,
    LongDashDotDot,
}

public enum XLLineStyle
{
    Single,
    ThinThin,
    ThinThick,
    ThickThin,
    ThickBetweenThin,
}

public interface IXLDrawingColorsAndLines
{
    public XLColor FillColor { get; set; }
    public double FillTransparency { get; set; }
    public XLColor LineColor { get; set; }
    public double LineTransparency { get; set; }
    public double LineWeight { get; set; }
    public XLDashStyle LineDash { get; set; }
    public XLLineStyle LineStyle { get; set; }

    public IXLDrawingStyle SetFillColor(XLColor value);
    public IXLDrawingStyle SetFillTransparency(double value);
    public IXLDrawingStyle SetLineColor(XLColor value);
    public IXLDrawingStyle SetLineTransparency(double value);
    public IXLDrawingStyle SetLineWeight(double value);
    public IXLDrawingStyle SetLineDash(XLDashStyle value);
    public IXLDrawingStyle SetLineStyle(XLLineStyle value);
}

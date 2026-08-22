#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFColorScaleMid
{
    public IXLCFColorScaleMax Midpoint(XLCFContentType type, string value, XLColor color);
    public IXLCFColorScaleMax Midpoint(XLCFContentType type, double value, XLColor color);
    public void Maximum(XLCFContentType type, string value, XLColor color);
    public void Maximum(XLCFContentType type, double value, XLColor color);
    public void HighestValue(XLColor color);
}

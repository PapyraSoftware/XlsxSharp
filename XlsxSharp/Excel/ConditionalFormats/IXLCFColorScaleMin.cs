#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public enum XLCFContentType
{
    Number,
    Percent,
    Formula,
    Percentile,
    Minimum,
    Maximum,
}

public interface IXLCFColorScaleMin
{
    public IXLCFColorScaleMid Minimum(XLCFContentType type, string value, XLColor color);
    public IXLCFColorScaleMid Minimum(XLCFContentType type, double value, XLColor color);
    public IXLCFColorScaleMid LowestValue(XLColor color);
}

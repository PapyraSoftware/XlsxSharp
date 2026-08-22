#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFColorScaleMax
{
    public void Maximum(XLCFContentType type, string value, XLColor color);
    public void Maximum(XLCFContentType type, double value, XLColor color);
    public void HighestValue(XLColor color);
}

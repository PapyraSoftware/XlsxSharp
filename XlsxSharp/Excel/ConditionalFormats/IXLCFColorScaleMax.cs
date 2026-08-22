#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFColorScaleMax
{
    void Maximum(XLCFContentType type, string value, XLColor color);
    void Maximum(XLCFContentType type, double value, XLColor color);
    void HighestValue(XLColor color);
}

#nullable disable

using System;

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFColorScaleMid
{
    IXLCFColorScaleMax Midpoint(XLCFContentType type, String value, XLColor color);
    IXLCFColorScaleMax Midpoint(XLCFContentType type, Double value, XLColor color);
    void Maximum(XLCFContentType type, String value, XLColor color);
    void Maximum(XLCFContentType type, Double value, XLColor color);
    void HighestValue(XLColor color);
}

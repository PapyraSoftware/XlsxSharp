#nullable disable

using System;

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFColorScaleMax
{
    void Maximum(XLCFContentType type, String value, XLColor color);
    void Maximum(XLCFContentType type, Double value, XLColor color);
    void HighestValue(XLColor color);
}

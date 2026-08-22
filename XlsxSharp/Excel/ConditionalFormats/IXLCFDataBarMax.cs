#nullable disable

using System;

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFDataBarMax
{
    void Maximum(XLCFContentType type, String value);
    void Maximum(XLCFContentType type, Double value);
    void HighestValue();
}

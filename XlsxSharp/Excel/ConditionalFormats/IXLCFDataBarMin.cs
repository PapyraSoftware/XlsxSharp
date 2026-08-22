#nullable disable

using System;

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFDataBarMin
{
    IXLCFDataBarMax Minimum(XLCFContentType type, string value);
    IXLCFDataBarMax Minimum(XLCFContentType type, double value);
    IXLCFDataBarMax LowestValue();
}

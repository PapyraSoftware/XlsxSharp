#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFDataBarMin
{
    public IXLCFDataBarMax Minimum(XLCFContentType type, string value);
    public IXLCFDataBarMax Minimum(XLCFContentType type, double value);
    public IXLCFDataBarMax LowestValue();
}

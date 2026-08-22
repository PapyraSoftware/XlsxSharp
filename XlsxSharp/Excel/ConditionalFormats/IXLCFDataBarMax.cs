#nullable disable

namespace XlsxSharp.Excel.ConditionalFormats;

public interface IXLCFDataBarMax
{
    public void Maximum(XLCFContentType type, string value);
    public void Maximum(XLCFContentType type, double value);
    public void HighestValue();
}

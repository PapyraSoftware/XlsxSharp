#nullable disable

using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFDataBarMin : IXLCFDataBarMin
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFDataBarMin(XLConditionalFormat conditionalFormat) =>
        this._conditionalFormat = conditionalFormat;

    public IXLCFDataBarMax Minimum(XLCFContentType type, string value)
    {
        this._conditionalFormat.ContentTypes.Initialize(type);
        this._conditionalFormat.Values.Initialize(new XLFormula { Value = value });
        return new XLCFDataBarMax(this._conditionalFormat);
    }

    public IXLCFDataBarMax Minimum(XLCFContentType type, double value) =>
        this.Minimum(type, value.ToInvariantString());

    public IXLCFDataBarMax LowestValue() => this.Minimum(XLCFContentType.Minimum, "0");
}

#nullable disable

using System;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFDataBarMin : IXLCFDataBarMin
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFDataBarMin(XLConditionalFormat conditionalFormat)
    {
        this._conditionalFormat = conditionalFormat;
    }

    public IXLCFDataBarMax Minimum(XLCFContentType type, String value)
    {
        this._conditionalFormat.ContentTypes.Initialize(type);
        this._conditionalFormat.Values.Initialize(new XLFormula { Value = value });
        return new XLCFDataBarMax(this._conditionalFormat);
    }

    public IXLCFDataBarMax Minimum(XLCFContentType type, Double value)
    {
        return this.Minimum(type, value.ToInvariantString());
    }

    public IXLCFDataBarMax LowestValue()
    {
        return this.Minimum(XLCFContentType.Minimum, "0");
    }
}

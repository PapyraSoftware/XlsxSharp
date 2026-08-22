#nullable disable

using System;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFDataBarMax : IXLCFDataBarMax
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFDataBarMax(XLConditionalFormat conditionalFormat)
    {
        this._conditionalFormat = conditionalFormat;
    }

    public void Maximum(XLCFContentType type, String value)
    {
        this._conditionalFormat.ContentTypes.Add(type);
        this._conditionalFormat.Values.Add(new XLFormula { Value = value });
    }

    public void Maximum(XLCFContentType type, Double value)
    {
        this.Maximum(type, value.ToInvariantString());
    }

    public void HighestValue()
    {
        this.Maximum(XLCFContentType.Maximum, "0");
    }
}

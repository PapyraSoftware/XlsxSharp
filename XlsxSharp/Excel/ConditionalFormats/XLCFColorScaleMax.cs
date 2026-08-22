#nullable disable

using System;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFColorScaleMax : IXLCFColorScaleMax
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFColorScaleMax(XLConditionalFormat conditionalFormat)
    {
        this._conditionalFormat = conditionalFormat;
    }

    public void Maximum(XLCFContentType type, String value, XLColor color)
    {
        this._conditionalFormat.Values.Add(new XLFormula { Value = value });
        this._conditionalFormat.Colors.Add(color);
        this._conditionalFormat.ContentTypes.Add(type);
    }

    public void Maximum(XLCFContentType type, Double value, XLColor color)
    {
        this.Maximum(type, value.ToInvariantString(), color);
    }

    public void HighestValue(XLColor color)
    {
        this._conditionalFormat.Values.Add(null);
        this._conditionalFormat.Colors.Add(color);
        this._conditionalFormat.ContentTypes.Add(XLCFContentType.Maximum);
    }
}

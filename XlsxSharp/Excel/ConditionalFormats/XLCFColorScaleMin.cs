#nullable disable

using System;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFColorScaleMin : IXLCFColorScaleMin
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFColorScaleMin(XLConditionalFormat conditionalFormat)
    {
        this._conditionalFormat = conditionalFormat;
    }

    public IXLCFColorScaleMid Minimum(XLCFContentType type, String value, XLColor color)
    {
        this._conditionalFormat.Values.Initialize(new XLFormula { Value = value });
        this._conditionalFormat.Colors.Initialize(color);
        this._conditionalFormat.ContentTypes.Initialize(type);
        return new XLCFColorScaleMid(this._conditionalFormat);
    }

    public IXLCFColorScaleMid Minimum(XLCFContentType type, Double value, XLColor color)
    {
        return this.Minimum(type, value.ToInvariantString(), color);
    }

    public IXLCFColorScaleMid LowestValue(XLColor color)
    {
        this._conditionalFormat.Values.Initialize(null);
        this._conditionalFormat.Colors.Initialize(color);
        this._conditionalFormat.ContentTypes.Initialize(XLCFContentType.Minimum);
        return new XLCFColorScaleMid(this._conditionalFormat);
    }
}

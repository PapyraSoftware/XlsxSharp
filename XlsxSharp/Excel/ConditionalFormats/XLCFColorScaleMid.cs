#nullable disable

using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFColorScaleMid : IXLCFColorScaleMid
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFColorScaleMid(XLConditionalFormat conditionalFormat) =>
        this._conditionalFormat = conditionalFormat;

    public IXLCFColorScaleMax Midpoint(XLCFContentType type, string value, XLColor color)
    {
        this._conditionalFormat.Values.Add(new XLFormula { Value = value });
        this._conditionalFormat.Colors.Add(color);
        this._conditionalFormat.ContentTypes.Add(type);
        return new XLCFColorScaleMax(this._conditionalFormat);
    }

    public IXLCFColorScaleMax Midpoint(XLCFContentType type, double value, XLColor color) =>
        this.Midpoint(type, value.ToInvariantString(), color);

    public void Maximum(XLCFContentType type, string value, XLColor color) =>
        this.Midpoint(type, value, color);

    public void Maximum(XLCFContentType type, double value, XLColor color) =>
        this.Maximum(type, value.ToInvariantString(), color);

    public void HighestValue(XLColor color)
    {
        this._conditionalFormat.Values.Initialize(null);
        this._conditionalFormat.Colors.Add(color);
        this._conditionalFormat.ContentTypes.Add(XLCFContentType.Maximum);
    }
}

#nullable disable

using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLCFIconSet : IXLCFIconSet
{
    private readonly XLConditionalFormat _conditionalFormat;

    public XLCFIconSet(XLConditionalFormat conditionalFormat) =>
        this._conditionalFormat = conditionalFormat;

    public IXLCFIconSet AddValue(
        XLCFIconSetOperator setOperator,
        string value,
        XLCFContentType type
    )
    {
        this._conditionalFormat.IconSetOperators.Add(setOperator);
        this._conditionalFormat.Values.Add(new XLFormula { Value = value });
        this._conditionalFormat.ContentTypes.Add(type);
        return new XLCFIconSet(this._conditionalFormat);
    }

    public IXLCFIconSet AddValue(
        XLCFIconSetOperator setOperator,
        double value,
        XLCFContentType type
    ) => this.AddValue(setOperator, value.ToInvariantString(), type);
}

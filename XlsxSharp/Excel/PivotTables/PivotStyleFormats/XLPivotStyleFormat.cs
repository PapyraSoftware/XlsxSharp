using System;

namespace XlsxSharp.Excel.PivotStyleFormats;

internal class XLPivotStyleFormat : XLPivotStyleFormatBase
{
    private readonly Func<XLPivotArea, bool> _filter;
    private readonly Func<XLPivotArea> _factory;

    public XLPivotStyleFormat(
        XLPivotTable pivotTable,
        Func<XLPivotArea, bool> filter,
        Func<XLPivotArea> factory
    )
        : base(pivotTable)
    {
        this._filter = filter;
        this._factory = factory;
    }

    internal override XLPivotArea GetCurrentArea()
    {
        return this._factory();
    }

    internal override bool Filter(XLPivotArea area)
    {
        return this._filter(area);
    }
}

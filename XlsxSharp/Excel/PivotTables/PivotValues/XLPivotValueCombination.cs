#nullable disable

namespace XlsxSharp.Excel.PivotValues;

internal class XLPivotValueCombination : IXLPivotValueCombination
{
    private readonly IXLPivotValue _pivotValue;

    public XLPivotValueCombination(IXLPivotValue pivotValue)
    {
        this._pivotValue = pivotValue;
    }

    public IXLPivotValue And(XLCellValue item)
    {
        return this
            ._pivotValue.SetBaseItemValue(item)
            .SetCalculationItem(XLPivotCalculationItem.Value);
    }

    public IXLPivotValue AndNext()
    {
        return this._pivotValue.SetCalculationItem(XLPivotCalculationItem.Next);
    }

    public IXLPivotValue AndPrevious()
    {
        return this._pivotValue.SetCalculationItem(XLPivotCalculationItem.Previous);
    }
}

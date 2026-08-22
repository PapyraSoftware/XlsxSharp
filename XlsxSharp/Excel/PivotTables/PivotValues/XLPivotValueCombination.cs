#nullable disable

namespace XlsxSharp.Excel.PivotValues;

internal class XLPivotValueCombination : IXLPivotValueCombination
{
    private readonly IXLPivotValue _pivotValue;

    public XLPivotValueCombination(IXLPivotValue pivotValue) => this._pivotValue = pivotValue;

    public IXLPivotValue And(XLCellValue item) =>
        this._pivotValue.SetBaseItemValue(item).SetCalculationItem(XLPivotCalculationItem.Value);

    public IXLPivotValue AndNext() =>
        this._pivotValue.SetCalculationItem(XLPivotCalculationItem.Next);

    public IXLPivotValue AndPrevious() =>
        this._pivotValue.SetCalculationItem(XLPivotCalculationItem.Previous);
}

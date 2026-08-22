// Keep this file CodeMaid organised and cleaned
namespace XlsxSharp.Excel.PivotStyleFormats;

internal class XLPivotTableStyleFormats : IXLPivotTableStyleFormats
{
    private readonly XLPivotTable _pivotTable;

    public XLPivotTableStyleFormats(XLPivotTable pivotTable)
    {
        this._pivotTable = pivotTable;
    }

    #region IXLPivotTableStyleFormats members

    public IXLPivotStyleFormats ColumnGrandTotalFormats =>
        new XLPivotStyleFormats(this._pivotTable, isRowGrand: false);

    public IXLPivotStyleFormats RowGrandTotalFormats =>
        new XLPivotStyleFormats(this._pivotTable, isRowGrand: true);

    #endregion IXLPivotTableStyleFormats members
}

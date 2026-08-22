namespace XlsxSharp.Excel;

internal class XLFilterConnector : IXLFilterConnector
{
    private readonly XLFilterColumn _filterColumn;

    public XLFilterConnector(XLFilterColumn filterColumn) => this._filterColumn = filterColumn;

    public IXLCustomFilteredColumn And =>
        new XLCustomFilteredColumn(this._filterColumn, XLConnector.And);

    public IXLCustomFilteredColumn Or =>
        new XLCustomFilteredColumn(this._filterColumn, XLConnector.Or);
}

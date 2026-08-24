#nullable disable

namespace XlsxSharp.Excel.CalcEngine;

internal class XLAddressComparer : IEqualityComparer<IXLAddress>
{
    private readonly bool _ignoreFixed;

    public XLAddressComparer(bool ignoreFixed) => this._ignoreFixed = ignoreFixed;

    public bool Equals(IXLAddress x, IXLAddress y) =>
        (x == null && y == null)
        || (
            x != null
            && y != null
            && string.Equals(
                x.Worksheet.Name,
                y.Worksheet.Name,
                StringComparison.InvariantCultureIgnoreCase
            )
            && x.ColumnNumber == y.ColumnNumber
            && x.RowNumber == y.RowNumber
            && (this._ignoreFixed || x.FixedColumn == y.FixedColumn && x.FixedRow == y.FixedRow)
        );

    public int GetHashCode(IXLAddress obj) =>
        new
        {
            WorksheetName = obj.Worksheet.Name.ToUpperInvariant(),
            obj.ColumnNumber,
            obj.RowNumber,
            FixedColumn = (this._ignoreFixed ? false : obj.FixedColumn),
            FixedRow = (this._ignoreFixed ? false : obj.FixedRow),
        }.GetHashCode();
}

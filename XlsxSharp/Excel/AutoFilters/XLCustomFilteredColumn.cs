using System;

namespace XlsxSharp.Excel;

internal class XLCustomFilteredColumn : IXLCustomFilteredColumn
{
    private readonly XLFilterColumn _filterColumn;
    private readonly XLConnector _connector;

    public XLCustomFilteredColumn(XLFilterColumn filterColumn, XLConnector connector)
    {
        this._filterColumn = filterColumn;
        this._connector = connector;
    }

    public void EqualTo(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.Equal, reapply);

    public void NotEqualTo(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.NotEqual, reapply);

    public void GreaterThan(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.GreaterThan, reapply);

    public void LessThan(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.LessThan, reapply);

    public void EqualOrGreaterThan(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);

    public void EqualOrLessThan(XLCellValue value, Boolean reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);

    public void BeginsWith(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter(value + "*", true, reapply);

    public void NotBeginsWith(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter(value + "*", false, reapply);

    public void EndsWith(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter("*" + value, true, reapply);

    public void NotEndsWith(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter("*" + value, false, reapply);

    public void Contains(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter("*" + value + "*", true, reapply);

    public void NotContains(String value, Boolean reapply) =>
        this.ApplyWildcardCustomFilter("*" + value + "*", false, reapply);

    private void ApplyCustomFilter(XLCellValue value, XLFilterOperator op, bool reapply) =>
        this._filterColumn.AddFilter(
            XLFilter.CreateCustomFilter(value, op, this._connector),
            reapply
        );

    private void ApplyWildcardCustomFilter(string pattern, bool match, bool reapply) =>
        this._filterColumn.AddFilter(
            XLFilter.CreateCustomPatternFilter(pattern, match, this._connector),
            reapply
        );
}

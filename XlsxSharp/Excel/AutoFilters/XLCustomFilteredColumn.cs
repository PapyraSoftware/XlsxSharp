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

    public void EqualTo(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.Equal, reapply);

    public void NotEqualTo(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.NotEqual, reapply);

    public void GreaterThan(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.GreaterThan, reapply);

    public void LessThan(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.LessThan, reapply);

    public void EqualOrGreaterThan(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);

    public void EqualOrLessThan(XLCellValue value, bool reapply) =>
        this.ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);

    public void BeginsWith(string value, bool reapply) =>
        this.ApplyWildcardCustomFilter(value + "*", true, reapply);

    public void NotBeginsWith(string value, bool reapply) =>
        this.ApplyWildcardCustomFilter(value + "*", false, reapply);

    public void EndsWith(string value, bool reapply) =>
        this.ApplyWildcardCustomFilter("*" + value, true, reapply);

    public void NotEndsWith(string value, bool reapply) =>
        this.ApplyWildcardCustomFilter("*" + value, false, reapply);

    public void Contains(string value, bool reapply) =>
        this.ApplyWildcardCustomFilter("*" + value + "*", true, reapply);

    public void NotContains(string value, bool reapply) =>
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

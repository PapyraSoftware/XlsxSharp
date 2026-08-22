namespace XlsxSharp.Excel;

public interface IXLCustomFilteredColumn
{
    public void EqualTo(XLCellValue value, bool reapply = true);
    public void NotEqualTo(XLCellValue value, bool reapply = true);
    public void GreaterThan(XLCellValue value, bool reapply = true);
    public void LessThan(XLCellValue value, bool reapply = true);
    public void EqualOrGreaterThan(XLCellValue value, bool reapply = true);
    public void EqualOrLessThan(XLCellValue value, bool reapply = true);
    public void BeginsWith(string value, bool reapply = true);
    public void NotBeginsWith(string value, bool reapply = true);
    public void EndsWith(string value, bool reapply = true);
    public void NotEndsWith(string value, bool reapply = true);
    public void Contains(string value, bool reapply = true);
    public void NotContains(string value, bool reapply = true);
}

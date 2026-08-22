namespace XlsxSharp.Excel;

public interface IXLFilterConnector
{
    public IXLCustomFilteredColumn And { get; }
    public IXLCustomFilteredColumn Or { get; }
}

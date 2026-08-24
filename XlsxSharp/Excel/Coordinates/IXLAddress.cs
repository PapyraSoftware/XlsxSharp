#nullable disable

namespace XlsxSharp.Excel;

/// <summary>
/// Reference to a single cell in a workbook. Reference can be absolute, relative or mixed.
/// Reference can be with or without a worksheet.
/// </summary>
public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>
{
    public string ColumnLetter { get; }
    public int ColumnNumber { get; }
    public bool FixedColumn { get; }
    public bool FixedRow { get; }
    public int RowNumber { get; }
    public string UniqueId { get; }

    /// <summary>
    /// Worksheet of the reference. Value is null for address without a worksheet.
    /// </summary>
    public IXLWorksheet Worksheet { get; }

    public string ToString(XLReferenceStyle referenceStyle);

    public string ToString(XLReferenceStyle referenceStyle, bool includeSheet);

    public string ToStringFixed();

    public string ToStringFixed(XLReferenceStyle referenceStyle);

    public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet);

    public string ToStringRelative();

    public string ToStringRelative(bool includeSheet);
}

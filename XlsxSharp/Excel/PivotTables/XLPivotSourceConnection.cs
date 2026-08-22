using System;

namespace XlsxSharp.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes data from a connection
/// to external source of data (e.g. database or a workbook).
/// </summary>
internal sealed class XLPivotSourceConnection : IXLPivotSource
{
    public XLPivotSourceConnection(uint connectionId) => this.ConnectionId = connectionId;

    public uint ConnectionId { get; }

    public bool Equals(IXLPivotSource otherSource)
    {
        XLPivotSourceConnection? other = otherSource as XLPivotSourceConnection;
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return this.ConnectionId == other.ConnectionId;
    }

    public override bool Equals(object? obj) => obj is IXLPivotSource other && this.Equals(other);

    public override int GetHashCode() => HashCode.Combine(this.ConnectionId).GetHashCode();

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea) =>
        throw new NotImplementedException(
            "Pivot cache source using a connection is not supported."
        );
}

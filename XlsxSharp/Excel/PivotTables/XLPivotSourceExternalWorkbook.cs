using System;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes data from external workbook.
/// </summary>
internal sealed class XLPivotSourceExternalWorkbook : IXLPivotSource
{
    /// <summary>
    /// External workbook relId. If relationships of cache definition changes, make sure to either keep same or update it.
    /// </summary>
    internal string RelId { get; }

    /// <summary>
    /// Are source data in external workbook defined by a <see cref="TableOrName"/> or by <see cref="Area">cell area</see>.
    /// </summary>
    [MemberNotNullWhen(true, nameof(TableOrName))]
    [MemberNotNullWhen(false, nameof(Area))]
    internal bool UsesName => this.TableOrName is not null;

    /// <summary>
    /// A table or defined name in an external workbook that contains source data.
    /// </summary>
    internal string? TableOrName { get; }

    /// <summary>
    /// An area in an external workbook that contains source data.
    /// </summary>
    internal SheetArea? Area { get; }

    internal XLPivotSourceExternalWorkbook(string relId, SheetArea area)
    {
        this.RelId = relId;
        this.Area = area;
    }

    internal XLPivotSourceExternalWorkbook(string relId, string tableOrName)
    {
        this.RelId = relId;
        this.TableOrName = tableOrName;
    }

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea)
    {
        throw new NotImplementedException("External workbook source is not supported.");
    }

    public bool Equals(IXLPivotSource otherSource)
    {
        XLPivotSourceExternalWorkbook? other = otherSource as XLPivotSourceExternalWorkbook;
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        // Two same RelIds could in theory point to different workbooks. I am not supporting
        // external sources for now anyway, so no unification through equality.
        return false;
    }

    public override bool Equals(object? other)
    {
        return other is IXLPivotSource source && this.Equals(source);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(this.RelId, this.Area).GetHashCode();
    }
}

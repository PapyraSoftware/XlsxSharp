using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

/// <summary>
/// A reference to the source data of <see cref="XLPivotCache"/>. The source might exist
/// or not, that is evaluated during pivot cache record refresh.
/// </summary>
internal sealed class XLPivotSourceReference : IXLPivotSource
{
    internal XLPivotSourceReference(SheetArea area)
    {
        this.Area = area;
        this.Name = null;
    }

    internal XLPivotSourceReference(string namedRangeOrTable)
    {
        this.Area = null;
        this.Name = namedRangeOrTable;
    }

    /// <summary>
    /// Are source data in external workbook defined by a <see cref="Name"/> or by <see cref="Area">cell area</see>.
    /// </summary>
    [MemberNotNullWhen(true, nameof(Name))]
    [MemberNotNullWhen(false, nameof(Area))]
    internal bool UsesName => this.Name is not null;

    /// <summary>
    /// Book area with the source data. Either this or <see cref="Name"/> is set.
    /// </summary>
    internal SheetArea? Area { get; }

    /// <summary>
    /// Name of a table or a book-scoped named range that contain the source data.
    /// Either this or <see cref="Area"/> is set.
    /// </summary>
    internal string? Name { get; }

    public bool Equals(IXLPivotSource otherSource)
    {
        XLPivotSourceReference? other = otherSource as XLPivotSourceReference;
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return Nullable.Equals(this.Area, other.Area)
            && XlsxSharp.XLHelper.NameComparer.Equals(this.Name, other.Name);
    }

    public override bool Equals(object? obj) => obj is IXLPivotSource other && this.Equals(other);

    public override int GetHashCode()
    {
        // Name is hashed through NameComparer so that it matches the case insensitive Equals.
        int nameHashCode = this.Name is not null
            ? XlsxSharp.XLHelper.NameComparer.GetHashCode(this.Name)
            : 0;
        return HashCode.Combine(this.Area, nameHashCode);
    }

    /// <summary>
    /// Try to determine actual area of the source reference in the
    /// workbook. Source reference might not be valid in the workbook.
    /// </summary>
    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea)
    {
        if (this.Name is not null)
        {
            // TODO: Named ranges are currently unusable, so only check tables.
            if (workbook.TryGetTable(this.Name, out XLTable table))
            {
                sheet = table.Worksheet;
                sheetArea = table.Area;
                return true;
            }

            sheet = null;
            sheetArea = null;
            return false;
        }

        Debug.Assert(this.Area is not null);
        if (workbook.WorksheetsInternal.TryGetWorksheet(this.Area.Value.Name, out sheet))
        {
            sheetArea = this.Area.Value.Area;
            return true;
        }

        sheetArea = default;
        return false;
    }
}

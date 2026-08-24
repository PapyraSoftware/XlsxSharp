using XlsxSharp.Extensions;
using XlsxSharp.Parser;

namespace XlsxSharp.Excel;

/// <summary>
/// A single point in a workbook. The book point might point to a deleted
/// worksheet, so it might be invalid. Make sure it is checked when
/// determining the properties of the actual data of the point.
/// </summary>
internal readonly struct SheetPoint : IEquatable<SheetPoint>
{
    internal SheetPoint(string sheetName, int row, int col)
        : this(sheetName, new Point(row, col)) { }

    internal SheetPoint(string sheetName, Point point)
    {
        ArgumentException.ThrowIfNullOrEmpty(sheetName);

        this.SheetName = sheetName;
        this.Point = point;
    }

    /// <summary>
    /// Name of the sheet. The sheet may be deleted.
    /// </summary>
    public string SheetName { get; }

    /// <inheritdoc cref="Excel.Point.Row"/>
    public int Row => this.Point.Column;

    /// <inheritdoc cref="Excel.Point.Column"/>
    public int Column => this.Point.Column;

    /// <summary>
    /// A point in the sheet.
    /// </summary>
    public Point Point { get; }

    public static bool operator ==(SheetPoint lhs, SheetPoint rhs) => lhs.Equals(rhs);

    public static bool operator !=(SheetPoint lhs, SheetPoint rhs) => !(lhs == rhs);

    public bool Equals(SheetPoint other) =>
        this.Point.Equals(other.Point)
        && XlsxSharp.XLHelper.SheetComparer.Equals(this.SheetName, other.SheetName);

    public override bool Equals(object? obj) => obj is SheetPoint other && this.Equals(other);

    // SheetName is hashed through SheetComparer so that it matches the case insensitive Equals.
    public override int GetHashCode() =>
        HashCode.Combine(XlsxSharp.XLHelper.SheetComparer.GetHashCode(this.SheetName), this.Point);

    public override string ToString()
    {
        string name = NameUtils.ShouldQuote(this.SheetName.AsSpan())
            ? this.SheetName.AlwaysEscapeSheetName()
            : this.SheetName;
        return $"{name}!{this.Point}";
    }
}

using System;
using System.Text;
using ClosedXML.Parser;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

/// <summary>
/// An immutable row address within a workbook.
/// </summary>
internal readonly record struct XLRowArea
{
    public XLRowArea(string name, int rowNumber)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);

        if (rowNumber is < XlsxSharp.XLHelper.MinRowNumber or > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(rowNumber));
        }

        this.Name = name;
        this.RowNumber = rowNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Row number, ranges from 1 to <see cref="XlsxSharp.XLHelper.MaxRowNumber"/>.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>
    /// Get the area of the row.
    /// </summary>
    public SheetArea Area =>
        new(
            this.Name,
            new Area(
                this.RowNumber,
                XlsxSharp.XLHelper.MinColumnNumber,
                this.RowNumber,
                XlsxSharp.XLHelper.MaxColumnNumber
            )
        );

    public bool Equals(XLRowArea other) =>
        this.RowNumber == other.RowNumber
        && XlsxSharp.XLHelper.SheetComparer.Equals(this.Name, other.Name);

    public override int GetHashCode()
    {
        unchecked
        {
            return (XlsxSharp.XLHelper.SheetComparer.GetHashCode(this.Name) * 397)
                ^ this.RowNumber.GetHashCode();
        }
    }

    public override string ToString()
    {
        string name = NameUtils.ShouldQuote(this.Name.AsSpan())
            ? this.Name.AlwaysEscapeSheetName()
            : this.Name;
        return new StringBuilder(name.Length + 1 + 7 + 1 + 7)
            .Append(name)
            .Append('!')
            .Append(this.RowNumber)
            .Append(':')
            .Append(this.RowNumber)
            .ToString();
    }
}

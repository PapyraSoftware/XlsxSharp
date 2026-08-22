using System;
using System.Text;
using ClosedXML.Parser;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

/// <summary>
/// An immutable column address within a workbook.
/// </summary>
internal readonly record struct XLColumnArea
{
    public XLColumnArea(string name, int columnNumber)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);

        if (
            columnNumber
            is < XlsxSharp.XLHelper.MinColumnNumber
                or > XlsxSharp.XLHelper.MaxColumnNumber
        )
        {
            throw new ArgumentOutOfRangeException(nameof(columnNumber));
        }

        this.Name = name;
        this.ColumNumber = columnNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Column number, ranges from 1 to <see cref="XlsxSharp.XLHelper.MaxColumnNumber"/>.
    /// </summary>
    public int ColumNumber { get; }

    public SheetArea Area =>
        new(
            this.Name,
            new Area(
                XlsxSharp.XLHelper.MinRowNumber,
                this.ColumNumber,
                XlsxSharp.XLHelper.MaxRowNumber,
                this.ColumNumber
            )
        );

    public bool Equals(XLColumnArea other) =>
        this.ColumNumber == other.ColumNumber
        && XlsxSharp.XLHelper.SheetComparer.Equals(this.Name, other.Name);

    // Name is hashed through SheetComparer so that it matches the case insensitive Equals.
    public override int GetHashCode() =>
        HashCode.Combine(XlsxSharp.XLHelper.SheetComparer.GetHashCode(this.Name), this.ColumNumber);

    public override string ToString()
    {
        string name = NameUtils.ShouldQuote(this.Name.AsSpan())
            ? this.Name.AlwaysEscapeSheetName()
            : this.Name;
        string column = XlsxSharp.XLHelper.GetColumnLetterFromNumber(this.ColumNumber);
        return new StringBuilder(name.Length + 2 + 2 * column.Length)
            .Append(name)
            .Append('!')
            .Append(column)
            .Append(':')
            .Append(column)
            .ToString();
    }
}

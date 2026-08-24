using System.Text;

namespace XlsxSharp.Parser;

/// <summary>
/// Due to frequency of an area in formulas, the grammar has a token that represents
/// an area in a sheet. This is the DTO from parser to engine. Two corners make an area
/// for A1 notation, but not for R1C1 (has several edge cases).
/// </summary>
public readonly record struct ReferenceArea
{
    /// <summary>
    /// First reference. First in terms of position in formula, not position
    /// in sheet.
    /// </summary>
    public RowCol First { get; }

    /// <summary>
    /// Second reference. Second in terms of position in formula, not position
    /// in sheet. If area was specified using only one cell, the value is
    /// same as <see cref="First"/>.
    /// </summary>
    public RowCol Second { get; }

    /// <summary>
    /// Semantic style of reference.
    /// </summary>
    public ReferenceStyle Style => this.First.Style;

    /// <summary>
    /// Is area a row span (e.g. <c>$5:7</c> in A1 or <c>R[7]</c>, <c>R7:R[9]</c>)?
    /// </summary>
    internal bool IsRowSpan => this.First.IsRow && this.Second.IsRow;

    /// <summary>
    /// Is area a col span (e.g. <c>$C:Z</c> in A1 or <c>C[7]</c>, <c>C7:C[9]</c>)?
    /// </summary>
    internal bool IsColSpan => this.First.IsColumn && this.Second.IsColumn;

    /// <summary>
    /// Create a reference symbol using the two <see cref="RowCol"/> (e.g.
    /// <c>A1:B2</c>) or two columns (e.g. <c>A:D</c>) or two rows (e.g.
    /// <c>7:8</c>).
    /// </summary>
    public ReferenceArea(RowCol first, RowCol second)
    {
        if (first.IsA1 ^ second.IsA1)
        {
            throw new ArgumentException("Both RowCol must use same semantic.");
        }

        this.First = first;
        this.Second = second;
    }

    /// <summary>
    /// Create an area for a single reference.
    /// </summary>
    public ReferenceArea(RowCol rowCol)
        : this(rowCol, rowCol)
    {
    }

    /// <summary>
    /// Create a new area from a single <see cref="RowCol"/>.
    /// </summary>
    /// <param name="rowType">Row axis type of a reference.</param>
    /// <param name="rowPosition">Row position.</param>
    /// <param name="columnType">Column axis type of a reference.</param>
    /// <param name="columnPosition">Column position.</param>
    /// <param name="style">Semantic of the reference.</param>
    public ReferenceArea(ReferenceAxisType rowType, int rowPosition, ReferenceAxisType columnType,
        int columnPosition, ReferenceStyle style)
        : this(new RowCol(rowType, rowPosition, columnType, columnPosition, style))
    {
    }

    /// <summary>
    /// Create a new area from a single row/column intersection.
    /// </summary>
    /// <param name="rowPosition"><see cref="ReferenceAxisType.Relative"/> row.</param>
    /// <param name="columnPosition"><see cref="ReferenceAxisType.Relative"/> column.</param>
    /// <param name="style">Semantic of the reference.</param>
    internal ReferenceArea(int rowPosition, int columnPosition, ReferenceStyle style)
        : this(new RowCol(ReferenceAxisType.Relative, rowPosition, ReferenceAxisType.Relative, columnPosition, style))
    {
    }

    /// <summary>
    /// Render area in A1 notation. The content must be a valid content
    /// from A1 token. If both references are same, only one is converted
    /// to display string.
    /// </summary>
    public string GetDisplayStringA1()
    {
        StringBuilder sb = new();
        this.AppendA1(sb);
        return sb.ToString();
    }

    /// <summary>
    /// Render area in R1C1 notation. The content must be a valid content
    /// from R1C1 token. If both references are same, only one is converted
    /// to display string.
    /// </summary>
    public string GetDisplayStringR1C1()
    {
        return this.AppendR1C1(new StringBuilder()).ToString();
    }

    /// <summary>
    /// Convert A1 reference to R1C1.
    /// </summary>
    /// <remarks>Assumes reference is in A1.</remarks>
    internal ReferenceArea ToR1C1(int anchorRow, int anchorCol)
    {
        RowCol first = this.First.ToR1C1(anchorRow, anchorCol);
        if (this.First != this.Second)
        {
            RowCol second = this.Second.ToR1C1(anchorRow, anchorCol);
            return new ReferenceArea(first, second);
        }

        return new ReferenceArea(first, first);
    }

    /// <summary>
    /// Convert R1C1 reference to A1.
    /// </summary>
    /// <remarks>Assumes reference is in R1C1.</remarks>
    internal ReferenceArea? ToA1OrError(int anchorRow, int anchorCol)
    {
        RowCol? first = this.First.ToA1OrError(anchorRow, anchorCol);
        if (first is null)
        {
            return null;
        }

        RowCol? second = this.First != this.Second ? this.Second.ToA1OrError(anchorRow, anchorCol) : first;
        if (second is null)
        {
            return null;
        }

        return new ReferenceArea(first.Value, second.Value);
    }

    internal StringBuilder Append(StringBuilder sb)
    {
        return this.Style == A1 ? this.AppendA1(sb) : this.AppendR1C1(sb);
    }

    /// <summary>
    /// Get reference in A1 notation.
    /// </summary>
    /// <remarks>Assumes reference is in A1.</remarks>
    internal StringBuilder AppendA1(StringBuilder sb)
    {
        // In A1, column and row span must have both parts specified.
        if (this.First == this.Second && !this.IsColSpan && !this.IsRowSpan)
        {
            return this.First.AppendA1(sb);
        }

        this.First.AppendA1(sb);
        sb.Append(':');
        this.Second.AppendA1(sb);
        return sb;
    }

    /// <summary>
    /// Get reference in R1C1 notation.
    /// </summary>
    /// <remarks>Assumes reference is in R1C1.</remarks>
    internal StringBuilder AppendR1C1(StringBuilder sb)
    {
        this.First.AppendR1C1(sb);

        if (this.First != this.Second)
        {
            sb.Append(':');
            this.Second.AppendR1C1(sb);
        }

        return sb;
    }
}

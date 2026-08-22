namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A border of differential format.
/// </summary>
/// <remarks>
/// Properties that are pivot/table only are often displayed only when pivot/table is refreshed.
/// Excel doesn't recalculate border styles for pivot table, unless necessary.
/// </remarks>
internal record XLDifferentialBorderValue
{
    internal static readonly XLDifferentialBorderValue Empty = new()
    {
        Left = null,
        Right = null,
        Top = null,
        Bottom = null,
        Diagonal = null,
        Vertical = null,
        Horizontal = null,
        DiagonalUp = false,
        DiagonalDown = false,
        Outline = true,
    };

    public required XLBorderLine? Left { get; init; }

    public required XLBorderLine? Right { get; init; }

    public required XLBorderLine? Top { get; init; }

    public required XLBorderLine? Bottom { get; init; }

    /// <summary>
    /// For pivot/tables only. Doesn't work on CF. Used when <see cref="DiagonalUp"/> or
    /// <see cref="DiagonalDown"/> are set. It's not possible to have different style
    /// for up/down diagonal.
    /// </summary>
    public required XLBorderLine? Diagonal { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required XLBorderLine? Vertical { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required XLBorderLine? Horizontal { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required bool DiagonalUp { get; init; }

    /// <summary>
    /// For pivot/tables only.
    /// </summary>
    public required bool DiagonalDown { get; init; }

    /// <summary>
    /// For pivot/tables or tables only.
    /// </summary>
    public required bool Outline { get; init; }
}

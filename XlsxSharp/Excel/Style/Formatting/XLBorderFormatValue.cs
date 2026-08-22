namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A border format master record.
/// </summary>
/// <remarks>
/// Many XML attributes are used only for dxf and thus are not represented in this structure (vertical, horizontal, outline).
/// </remarks>
internal record XLBorderFormatValue
{
    internal static readonly XLBorderFormatValue None = new()
    {
        Left = XLBorderLine.None,
        Right = XLBorderLine.None,
        Top = XLBorderLine.None,
        Bottom = XLBorderLine.None,
        Diagonal = XLBorderLine.None,
        DiagonalUp = false,
        DiagonalDown = false,
    };

    public required XLBorderLine Left { get; init; }

    public required XLBorderLine Right { get; init; }

    public required XLBorderLine Top { get; init; }

    public required XLBorderLine Bottom { get; init; }

    /// <summary>
    /// Used when <see cref="DiagonalUp"/> or <see cref="DiagonalDown"/> are set. It's not possible
    /// to have different style for up/down diagonal.
    /// </summary>
    public required XLBorderLine Diagonal { get; init; }

    public required bool DiagonalUp { get; init; }

    public required bool DiagonalDown { get; init; }

    internal static XLBorderFormatValue FromDxf(XLDifferentialBorderValue dxfBorder) =>
        new()
        {
            Left = dxfBorder.Left ?? XLBorderLine.None,
            Right = dxfBorder.Right ?? XLBorderLine.None,
            Top = dxfBorder.Top ?? XLBorderLine.None,
            Bottom = dxfBorder.Bottom ?? XLBorderLine.None,
            Diagonal = dxfBorder.Diagonal ?? XLBorderLine.None,
            DiagonalUp = dxfBorder.DiagonalUp,
            DiagonalDown = dxfBorder.DiagonalDown,
        };
}

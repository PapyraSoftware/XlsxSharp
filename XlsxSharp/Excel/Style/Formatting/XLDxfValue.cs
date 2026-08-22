namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A differential format.
/// </summary>
internal record XLDxfValue
{
    internal static readonly XLDxfValue Empty = new()
    {
        NumberFormat = null,
        Font = XLDifferentialFontValue.Empty,
        Fill = XLDifferentialFillValue.Empty,
        Alignment = XLDifferentialAlignmentValue.Empty,
        Border = XLDifferentialBorderValue.Empty,
        Protection = XLDifferentialProtectionValue.Empty,
    };

    public required string? NumberFormat { get; init; }

    public required XLDifferentialFontValue Font { get; init; }

    public required XLDifferentialFillValue Fill { get; init; }

    public required XLDifferentialAlignmentValue Alignment { get; init; }

    public required XLDifferentialBorderValue Border { get; init; }

    public required XLDifferentialProtectionValue Protection { get; init; }
}

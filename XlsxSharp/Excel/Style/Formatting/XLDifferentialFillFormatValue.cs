namespace XlsxSharp.Excel.Formatting;

internal record XLDifferentialFillValue
{
    internal static XLDifferentialFillValue Empty { get; } = new(null, null, null);

    internal XLDifferentialFillValue(XLPatternFill pattern)
        : this(pattern, null, null) { }

    internal XLDifferentialFillValue(XLFillFormatValue fill)
        : this(fill.Pattern, fill.LinearGradient, fill.PathGradient) { }

    private XLDifferentialFillValue(
        XLPatternFill? pattern,
        XLLinearGradientFill? linearGradient,
        XLPathGradientFill? pathGradient
    )
    {
        this.Pattern = pattern;
        this.LinearGradient = linearGradient;
        this.PathGradient = pathGradient;
    }

    internal XLPatternFill? Pattern { get; }

    internal XLLinearGradientFill? LinearGradient { get; }

    internal XLPathGradientFill? PathGradient { get; }
}

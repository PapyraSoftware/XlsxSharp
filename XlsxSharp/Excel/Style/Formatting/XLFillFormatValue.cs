namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A fill formatting. Fill can have at most one pattern that specifies how to fill a cell.
/// </summary>
internal record XLFillFormatValue
{
    internal static readonly XLFillFormatValue Empty = new(null, null, null);

    internal static readonly XLFillFormatValue None = new(
        new XLPatternFill
        {
            PatternType = XLFillPatternValues.None,
            BackgroundColor = XLColor.Automatic,
            PatternColor = XLColor.Automatic,
        }
    );

    internal static readonly XLFillFormatValue Gray125 = new(
        new XLPatternFill
        {
            PatternType = XLFillPatternValues.Gray125,
            BackgroundColor = XLColor.Automatic,
            PatternColor = XLColor.Automatic,
        }
    );

    public XLFillFormatValue(XLPatternFill pattern)
        : this(pattern, null, null) { }

    public XLFillFormatValue(XLLinearGradientFill linearGradient)
        : this(null, linearGradient, null) { }

    public XLFillFormatValue(XLPathGradientFill pathGradient)
        : this(null, null, pathGradient) { }

    private XLFillFormatValue(
        XLPatternFill? pattern,
        XLLinearGradientFill? linearGradient,
        XLPathGradientFill? pathGradient
    )
    {
        this.Pattern = pattern;
        this.LinearGradient = linearGradient;
        this.PathGradient = pathGradient;
    }

    public XLPatternFill? Pattern { get; }

    public XLLinearGradientFill? LinearGradient { get; }

    public XLPathGradientFill? PathGradient { get; }
}

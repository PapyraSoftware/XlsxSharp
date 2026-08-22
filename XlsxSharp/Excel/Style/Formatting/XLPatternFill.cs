namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A pattern fill format.
/// </summary>
internal record XLPatternFill
{
    /// <summary>
    /// Foreground color of the pattern (i.e., if pattern is a grid, this is the color to draw
    /// lines of the grid).
    /// </summary>
    internal required XLColor PatternColor { get; init; }

    /// <summary>
    /// Background color of the pattern.
    /// </summary>
    internal required XLColor BackgroundColor { get; init; }

    /// <summary>
    /// Shape of the pattern.
    /// </summary>
    internal required XLFillPatternValues PatternType { get; init; }

    internal XLPatternFill WithModifiedPattern(XLFillPatternValues patternType)
    {
        if (this.PatternType == XLFillPatternValues.None && patternType != XLFillPatternValues.None)
        {
            // Keep the original behavior. Just having a non-none pattern type doesn't actually
            // affects color of cells visually, it needs at least background color.
            return this with
            {
                BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1),
                PatternType = patternType,
            };
        }

        return this with
        {
            PatternType = patternType,
        };
    }

    internal XLPatternFill WithModifiedBgColor(XLColor newBgColor)
    {
        XLFillPatternValues newPattern;
        if (newBgColor.IsAuto && this.PatternType == XLFillPatternValues.Solid)
        {
            // Fill is single color, and that color will be is transparent -> there is no fill.
            newPattern = XLFillPatternValues.None;
        }
        else if (!newBgColor.IsAuto && this.PatternType == XLFillPatternValues.None)
        {
            // Color is not transparent, but fill is none -> make it solid
            newPattern = XLFillPatternValues.Solid;
        }
        else
        {
            newPattern = this.PatternType;
        }

        return new XLPatternFill
        {
            PatternType = newPattern,
            PatternColor = this.PatternColor,
            BackgroundColor = newBgColor,
        };
    }
}

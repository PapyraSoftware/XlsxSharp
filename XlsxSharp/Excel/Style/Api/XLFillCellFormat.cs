using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// An API object to modify fill of a cell range.
/// </summary>
/// <remarks>
/// Most of the various logic that adjusts background color or pattern based on changes to pattern
/// type or background color have been inherited from original implementation.
/// </remarks>
internal sealed partial class XLFillCellFormat
{
    private static readonly XLPatternFill PatternNone = new()
    {
        PatternType = XLFillPatternValues.None,
        PatternColor = XLColor.Automatic,
        BackgroundColor = XLColor.Automatic,
    };

    private readonly XLCellFormat _parent;

    public XLFillCellFormat(XLCellFormat parent)
    {
        this._parent = parent;
    }

    private XLColor BackgroundColor
    {
        get => this._parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).BackgroundColor);
        set =>
            this._parent.ModifyFill(
                static (fill, bgColor) =>
                {
                    XLPatternFill currentPattern = fill.Pattern ?? PatternNone;
                    XLPatternFill newPattern = currentPattern.WithModifiedBgColor(bgColor);
                    return new XLFillFormatValue(newPattern);
                },
                value
            );
    }

    private XLColor PatternColor
    {
        get => this._parent.Resolve(x => (x.Fill.Pattern ?? PatternNone).PatternColor);
        set =>
            this._parent.ModifyFill(
                (fill, patternColor) =>
                {
                    XLPatternFill pattern = fill.Pattern ?? PatternNone;
                    return new XLFillFormatValue(pattern with { PatternColor = patternColor });
                },
                value
            );
    }

    private XLFillPatternValues PatternType
    {
        get => this._parent.Resolve(static x => (x.Fill.Pattern ?? PatternNone).PatternType);
        set =>
            this._parent.ModifyFill(
                static (fill, patternType) =>
                {
                    XLPatternFill pattern = fill.Pattern ?? PatternNone;
                    XLPatternFill newPattern = pattern.WithModifiedPattern(patternType);
                    return new XLFillFormatValue(newPattern);
                },
                value
            );
    }

    public override bool Equals(object? obj)
    {
        return obj is IXLFill other && (this as IEquatable<IXLFill>).Equals(other);
    }

    public override int GetHashCode()
    {
        return 0;
    }

    internal void SetValue(IXLFill value)
    {
        // No need for shenanigans with changing pattern fill or colors, the original should be valid and consistent.
        this._parent.ModifyFill(
            static (_, patternFill) =>
                new XLFillFormatValue(
                    new XLPatternFill
                    {
                        PatternType = patternFill.PatternType,
                        PatternColor = patternFill.PatternColor,
                        BackgroundColor = patternFill.BackgroundColor,
                    }
                ),
            value
        );
    }
}

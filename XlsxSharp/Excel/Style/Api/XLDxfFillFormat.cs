using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// API object for dxf hidden behind IXLStyle.IXLFill interface.
/// </summary>
internal class XLDxfFillFormat : IXLFill
{
    private static readonly XLPatternFill _default = new()
    {
        PatternType = XLFillPatternValues.None,
        BackgroundColor = XLColor.Automatic,
        PatternColor = XLColor.Automatic,
    };

    private readonly XLDxFormat _parent;

    internal XLDxfFillFormat(XLDxFormat parent) => this._parent = parent;

    XLColor IXLFill.BackgroundColor
    {
        get => this.Resolve(static fill => fill.Pattern?.BackgroundColor, _default.BackgroundColor);
        set =>
            this.Modify(
                static (fill, bgColor) =>
                {
                    XLPatternFill currentPattern = fill.Pattern ?? _default;
                    XLPatternFill newPattern = currentPattern.WithModifiedBgColor(bgColor);
                    return new XLDifferentialFillValue(newPattern);
                },
                value
            );
    }

    XLColor IXLFill.PatternColor
    {
        get => this.Resolve(static fill => fill.Pattern?.PatternColor, _default.PatternColor);
        set =>
            this.Modify(
                static (fill, patternColor) =>
                {
                    XLPatternFill pattern = fill.Pattern ?? _default;
                    return new XLDifferentialFillValue(
                        pattern with
                        {
                            PatternColor = patternColor,
                        }
                    );
                },
                value
            );
    }

    XLFillPatternValues IXLFill.PatternType
    {
        get => this.Resolve(static fill => fill.Pattern?.PatternType, _default.PatternType);
        set =>
            this.Modify(
                static (fill, patternType) =>
                {
                    XLPatternFill pattern = fill.Pattern ?? _default;
                    XLPatternFill newPattern = pattern.WithModifiedPattern(patternType);
                    return new XLDifferentialFillValue(newPattern);
                },
                value
            );
    }

    IXLStyle IXLFill.SetBackgroundColor(XLColor value)
    {
        (this as IXLFill).BackgroundColor = value;
        return this._parent;
    }

    IXLStyle IXLFill.SetPatternColor(XLColor value)
    {
        (this as IXLFill).PatternColor = value;
        return this._parent;
    }

    IXLStyle IXLFill.SetPatternType(XLFillPatternValues value)
    {
        (this as IXLFill).PatternType = value;
        return this._parent;
    }

    bool IEquatable<IXLFill>.Equals(IXLFill other) => throw new NotSupportedException();

    internal void SetValue(IXLFill value) =>
        // The original should be valid and consistent.
        this.Modify(
            static (_, patternFill) =>
                new XLDifferentialFillValue(
                    new XLPatternFill
                    {
                        PatternType = patternFill.PatternType,
                        PatternColor = patternFill.PatternColor,
                        BackgroundColor = patternFill.BackgroundColor,
                    }
                ),
            value
        );

    private T Resolve<T>(Func<XLDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : struct =>
        this._parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;

    private T Resolve<T>(Func<XLDifferentialFillValue, T?> getProperty, T defaultValue)
        where T : class =>
        this._parent.Resolve(static format => format.Fill, getProperty) ?? defaultValue;

    private void Modify<T>(
        Func<XLDifferentialFillValue, T, XLDifferentialFillValue> modify,
        T value
    ) => this._parent.ModifyFill(modify, value);
}

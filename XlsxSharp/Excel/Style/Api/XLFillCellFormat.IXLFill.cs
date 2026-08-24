namespace XlsxSharp.Excel;

internal sealed partial class XLFillCellFormat : IXLFill
{
    XLColor IXLFill.BackgroundColor
    {
        get => this.BackgroundColor;
        set => this.BackgroundColor = value;
    }

    XLColor IXLFill.PatternColor
    {
        get => this.PatternColor;
        set => this.PatternColor = value;
    }

    XLFillPatternValues IXLFill.PatternType
    {
        get => this.PatternType;
        set => this.PatternType = value;
    }

    IXLStyle IXLFill.SetBackgroundColor(XLColor value)
    {
        this.BackgroundColor = value;
        return this._parent;
    }

    IXLStyle IXLFill.SetPatternColor(XLColor value)
    {
        this.PatternColor = value;
        return this._parent;
    }

    IXLStyle IXLFill.SetPatternType(XLFillPatternValues value)
    {
        this.PatternType = value;
        return this._parent;
    }

    bool IEquatable<IXLFill>.Equals(IXLFill other)
    {
        // This is a "business" equality, i.e. will both fills look the same.
        // This is gradient fill, other can only represent pattern, regardless of what it actually is.
        bool isPatternFill = this._parent.Resolve(static x => x.Fill.Pattern) is not null;
        if (!isPatternFill)
        {
            return false;
        }

        if (!HasFill(this) && !HasFill(other))
        {
            return true;
        }

        if (this.PatternType != other.PatternType)
        {
            return false;
        }

        if (this.BackgroundColor != other.BackgroundColor)
        {
            return false;
        }

        XLColor? patternColor = UsesPatternColor(this) ? this.PatternColor : null;
        XLColor? otherPatternColor = UsesPatternColor(other) ? other.PatternColor : null;
        if (patternColor != otherPatternColor)
        {
            return false;
        }

        return true;

        static bool HasFill(IXLFill fill)
        {
            XLFillPatternValues patternType = fill.PatternType;
            if (patternType == XLFillPatternValues.None)
            {
                return false;
            }

            if (patternType == XLFillPatternValues.Solid && fill.BackgroundColor.IsAuto)
            {
                return false;
            }

            return true;
        }

        static bool UsesPatternColor(IXLFill fill)
        {
            return fill.PatternType
                is not XLFillPatternValues.None
                    and not XLFillPatternValues.Solid;
        }
    }
}

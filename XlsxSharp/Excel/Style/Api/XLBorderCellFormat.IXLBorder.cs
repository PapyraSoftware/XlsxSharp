using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal sealed partial class XLBorderCellFormat : IXLBorder
{
    XLBorderStyleValues IXLBorder.OutsideBorder
    {
        set =>
            this._parent.ModifyOuterBorder(
                static (borderLine, style) => borderLine with { Style = style },
                value
            );
    }

    XLColor IXLBorder.OutsideBorderColor
    {
        set =>
            this._parent.ModifyOuterBorder(
                static (borderLine, color) => borderLine with { Color = color },
                value
            );
    }

    XLBorderStyleValues IXLBorder.InsideBorder
    {
        set =>
            this._parent.ModifyInnerBorder(
                static (borderLine, style) => borderLine with { Style = style },
                value
            );
    }

    XLColor IXLBorder.InsideBorderColor
    {
        set =>
            this._parent.ModifyInnerBorder(
                static (borderLine, color) => borderLine with { Color = color },
                value
            );
    }

    XLBorderStyleValues IXLBorder.LeftBorder
    {
        get => this.LeftBorder;
        set => this.LeftBorder = value;
    }

    XLColor IXLBorder.LeftBorderColor
    {
        get => this.LeftBorderColor;
        set => this.LeftBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.RightBorder
    {
        get => this.RightBorder;
        set => this.RightBorder = value;
    }

    XLColor IXLBorder.RightBorderColor
    {
        get => this.RightBorderColor;
        set => this.RightBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.TopBorder
    {
        get => this.TopBorder;
        set => this.TopBorder = value;
    }

    XLColor IXLBorder.TopBorderColor
    {
        get => this.TopBorderColor;
        set => this.TopBorderColor = value;
    }

    XLBorderStyleValues IXLBorder.BottomBorder
    {
        get => this.BottomBorder;
        set => this.BottomBorder = value;
    }

    XLColor IXLBorder.BottomBorderColor
    {
        get => this.BottomBorderColor;
        set => this.BottomBorderColor = value;
    }

    bool IXLBorder.DiagonalUp
    {
        get => this.DiagonalUp;
        set => this.DiagonalUp = value;
    }

    bool IXLBorder.DiagonalDown
    {
        get => this.DiagonalDown;
        set => this.DiagonalDown = value;
    }

    XLBorderStyleValues IXLBorder.DiagonalBorder
    {
        get => this.DiagonalBorder;
        set => this.DiagonalBorder = value;
    }

    XLColor IXLBorder.DiagonalBorderColor
    {
        get => this.DiagonalBorderColor;
        set => this.DiagonalBorderColor = value;
    }

    IXLStyle IXLBorder.SetOutsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).OutsideBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetOutsideBorderColor(XLColor value)
    {
        (this as IXLBorder).OutsideBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetInsideBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).InsideBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetInsideBorderColor(XLColor value)
    {
        (this as IXLBorder).InsideBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetLeftBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).LeftBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetLeftBorderColor(XLColor value)
    {
        (this as IXLBorder).LeftBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetRightBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).RightBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetRightBorderColor(XLColor value)
    {
        (this as IXLBorder).RightBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetTopBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).TopBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetTopBorderColor(XLColor value)
    {
        (this as IXLBorder).TopBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetBottomBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).BottomBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetBottomBorderColor(XLColor value)
    {
        (this as IXLBorder).BottomBorderColor = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetDiagonalUp() => (this as IXLBorder).SetDiagonalUp(true);

    IXLStyle IXLBorder.SetDiagonalUp(bool value)
    {
        (this as IXLBorder).DiagonalUp = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetDiagonalDown() => (this as IXLBorder).SetDiagonalDown(true);

    IXLStyle IXLBorder.SetDiagonalDown(bool value)
    {
        (this as IXLBorder).DiagonalDown = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorder(XLBorderStyleValues value)
    {
        (this as IXLBorder).DiagonalBorder = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetDiagonalBorderColor(XLColor value)
    {
        (this as IXLBorder).DiagonalBorderColor = value;
        return this._parent;
    }

    bool IEquatable<IXLBorder>.Equals(IXLBorder? other)
    {
        // A "business" equals, when borders are visually same, they are considered equals.
        if (other is null)
        {
            return false;
        }

        XLBorderFormatValue thisBorder = this._parent.Resolve(static x => x.Border);
        XLBorderLine otherLeft = new(other.LeftBorderColor, other.LeftBorder);
        if (!IsSameLine(thisBorder.Left, otherLeft))
        {
            return false;
        }

        XLBorderLine otherTop = new(other.TopBorderColor, other.TopBorder);
        if (!IsSameLine(thisBorder.Top, otherTop))
        {
            return false;
        }

        XLBorderLine otherRight = new(other.RightBorderColor, other.RightBorder);
        if (!IsSameLine(thisBorder.Right, otherRight))
        {
            return false;
        }

        XLBorderLine otherBottom = new(other.BottomBorderColor, other.BottomBorder);
        if (!IsSameLine(thisBorder.Bottom, otherBottom))
        {
            return false;
        }

        // Check diagonals. If the diagonal flag is not set, the diagonal is not displayed. Normalize them to take into the account the direction flag
        XLBorderLine thisDiagonalUp = MakeThisDiagonal(thisBorder.Diagonal, thisBorder.DiagonalUp);
        XLBorderLine otherDiagonalUp = MakeOtherDiagonal(other, other.DiagonalUp);
        if (!IsSameLine(thisDiagonalUp, otherDiagonalUp))
        {
            return false;
        }

        XLBorderLine thisDiagonalDown = MakeThisDiagonal(
            thisBorder.Diagonal,
            thisBorder.DiagonalDown
        );
        XLBorderLine otherDiagonalDown = MakeOtherDiagonal(other, other.DiagonalDown);
        if (!IsSameLine(thisDiagonalDown, otherDiagonalDown))
        {
            return false;
        }

        return true;

        static bool IsSameLine(XLBorderLine lhs, XLBorderLine rhs)
        {
            if (lhs.Style == XLBorderStyleValues.None && rhs.Style == XLBorderStyleValues.None)
            {
                return true;
            }

            // Auto color in context of border is black, not transparent.
            return lhs.Style == rhs.Style && lhs.Color == rhs.Color;
        }

        static XLBorderLine MakeThisDiagonal(XLBorderLine diagonal, bool diagonalDirection)
        {
            return diagonal with
            {
                Style = diagonalDirection ? diagonal.Style : XLBorderStyleValues.None,
            };
        }

        static XLBorderLine MakeOtherDiagonal(IXLBorder border, bool diagonalDirection)
        {
            return new XLBorderLine(
                border.DiagonalBorderColor,
                diagonalDirection ? border.DiagonalBorder : XLBorderStyleValues.None
            );
        }
    }
}

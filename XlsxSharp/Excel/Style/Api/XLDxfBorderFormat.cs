using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

// TODO Styles: Plan is to have a separate interface for dxf. Implement necessary features to pass tests, then replace.
/// <summary>
/// API object for dxf hidden behind IXLStyle.IXLBorder interface.
/// </summary>
internal class XLDxfBorderFormat : IXLBorder
{
    private static readonly XLBorderLine NoLine = XLBorderLine.None;
    private readonly XLDxFormat _parent;

    internal XLDxfBorderFormat(XLDxFormat parent)
    {
        this._parent = parent;
    }

    XLBorderStyleValues IXLBorder.OutsideBorder
    {
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Outline = true,
                        Left = (border.Left ?? NoLine) with { Style = style },
                        Right = (border.Right ?? NoLine) with { Style = style },
                        Top = (border.Top ?? NoLine) with { Style = style },
                        Bottom = (border.Bottom ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.OutsideBorderColor
    {
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Outline = true,
                        Left = (border.Left ?? NoLine) with { Color = color },
                        Right = (border.Right ?? NoLine) with { Color = color },
                        Top = (border.Top ?? NoLine) with { Color = color },
                        Bottom = (border.Bottom ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    XLBorderStyleValues IXLBorder.InsideBorder
    {
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Outline = false,
                        Left = (border.Left ?? NoLine) with { Style = style },
                        Right = (border.Right ?? NoLine) with { Style = style },
                        Top = (border.Top ?? NoLine) with { Style = style },
                        Bottom = (border.Bottom ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.InsideBorderColor
    {
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Outline = false,
                        Left = (border.Left ?? NoLine) with { Color = color },
                        Right = (border.Right ?? NoLine) with { Color = color },
                        Top = (border.Top ?? NoLine) with { Color = color },
                        Bottom = (border.Bottom ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    XLBorderStyleValues IXLBorder.LeftBorder
    {
        get => this.Resolve(static x => x.Left?.Style, NoLine.Style);
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Left = (border.Left ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.LeftBorderColor
    {
        get => this.Resolve(static border => border.Left?.Color, NoLine.Color);
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Left = (border.Left ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    XLBorderStyleValues IXLBorder.RightBorder
    {
        get => this.Resolve(static x => x.Right?.Style, NoLine.Style);
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Right = (border.Right ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.RightBorderColor
    {
        get => this.Resolve(static border => border.Right?.Color, NoLine.Color);
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Right = (border.Right ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    XLBorderStyleValues IXLBorder.TopBorder
    {
        get => this.Resolve(static x => x.Top?.Style, NoLine.Style);
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Top = (border.Top ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.TopBorderColor
    {
        get => this.Resolve(static border => border.Top?.Color, NoLine.Color);
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Top = (border.Top ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    XLBorderStyleValues IXLBorder.BottomBorder
    {
        get => this.Resolve(static x => x.Bottom?.Style, NoLine.Style);
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Bottom = (border.Bottom ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.BottomBorderColor
    {
        get => this.Resolve(static border => border.Bottom?.Color, NoLine.Color);
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Bottom = (border.Bottom ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    bool IXLBorder.DiagonalUp
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    bool IXLBorder.DiagonalDown
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    XLBorderStyleValues IXLBorder.DiagonalBorder
    {
        get => this.Resolve(static x => x.Diagonal?.Style, NoLine.Style);
        set =>
            this.Modify(
                static (border, style) =>
                    border with
                    {
                        Diagonal = (border.Diagonal ?? NoLine) with { Style = style },
                    },
                value
            );
    }

    XLColor IXLBorder.DiagonalBorderColor
    {
        get => this.Resolve(static border => border.Diagonal?.Color, NoLine.Color);
        set =>
            this.Modify(
                static (border, color) =>
                    border with
                    {
                        Diagonal = (border.Diagonal ?? NoLine) with { Color = color },
                    },
                value
            );
    }

    internal void SetValue(IXLBorder value)
    {
        this.Modify(
            (border, other) =>
                border with
                {
                    Left = new XLBorderLine(other.LeftBorderColor, other.LeftBorder),
                    Right = new XLBorderLine(other.RightBorderColor, other.RightBorder),
                    Top = new XLBorderLine(other.TopBorderColor, other.TopBorder),
                    Bottom = new XLBorderLine(other.BottomBorderColor, other.BottomBorder),
                    DiagonalUp = other.DiagonalUp,
                    DiagonalDown = other.DiagonalDown,
                    Diagonal = new XLBorderLine(other.DiagonalBorderColor, other.DiagonalBorder),
                    Horizontal = null,
                    Vertical = null,
                    Outline = true,
                },
            value
        );
    }

    private T Resolve<T>(Func<XLDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return this._parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private T Resolve<T>(Func<XLDifferentialBorderValue, T?> getProperty, T defaultValue)
        where T : class
    {
        return this._parent.Resolve(static format => format.Border, getProperty) ?? defaultValue;
    }

    private void Modify<T>(
        Func<XLDifferentialBorderValue, T, XLDifferentialBorderValue> modify,
        T value
    )
    {
        this._parent.ModifyBorder(modify, value);
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

    IXLStyle IXLBorder.SetDiagonalUp()
    {
        return (this as IXLBorder).SetDiagonalUp(true);
    }

    IXLStyle IXLBorder.SetDiagonalUp(bool value)
    {
        (this as IXLBorder).DiagonalUp = value;
        return this._parent;
    }

    IXLStyle IXLBorder.SetDiagonalDown()
    {
        return (this as IXLBorder).SetDiagonalDown(true);
    }

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

    bool IEquatable<IXLBorder>.Equals(IXLBorder other)
    {
        throw new NotImplementedException();
    }
}

using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal class XLDxfAlignmentFormat : IXLAlignment
{
    private readonly XLAlignmentFormatValue _default = XLAlignmentFormatValue.Default;
    private readonly XLDxFormat _parent;

    internal XLDxfAlignmentFormat(XLDxFormat parent) => this._parent = parent;

    XLAlignmentHorizontalValues IXLAlignment.Horizontal
    {
        get => this.Resolve(static alignment => alignment.Horizontal, this._default.Horizontal);
        set =>
            this.Modify(
                static (alignment, horizontal) => alignment with { Horizontal = horizontal },
                value
            );
    }

    XLAlignmentVerticalValues IXLAlignment.Vertical
    {
        get => this.Resolve(static alignment => alignment.Vertical, this._default.Vertical);
        set =>
            this.Modify(
                static (alignment, vertical) => alignment with { Vertical = vertical },
                value
            );
    }

    int IXLAlignment.Indent
    {
        get => this.Resolve(static alignment => alignment.Indent, this._default.Indent);
        set => this.Modify(static (alignment, indent) => alignment with { Indent = indent }, value);
    }

    bool IXLAlignment.JustifyLastLine
    {
        get =>
            this.Resolve(
                static alignment => alignment.JustifyLastLine,
                this._default.JustifyLastLine
            );
        set =>
            this.Modify(
                static (alignment, justifyLastLine) =>
                    alignment with
                    {
                        JustifyLastLine = justifyLastLine,
                    },
                value
            );
    }

    XLAlignmentReadingOrderValues IXLAlignment.ReadingOrder
    {
        get => this.Resolve(static alignment => alignment.ReadingOrder, this._default.ReadingOrder);
        set =>
            this.Modify(
                static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder },
                value
            );
    }

    int IXLAlignment.RelativeIndent
    {
        get =>
            this.Resolve(
                static alignment => alignment.RelativeIndent,
                this._default.RelativeIndent
            );
        set =>
            this.Modify(
                static (alignment, relativeIndent) =>
                    alignment with
                    {
                        RelativeIndent = relativeIndent,
                    },
                value
            );
    }

    bool IXLAlignment.ShrinkToFit
    {
        get => this.Resolve(static alignment => alignment.ShrinkToFit, this._default.ShrinkToFit);
        set =>
            this.Modify(
                static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit },
                value
            );
    }

    int IXLAlignment.TextRotation
    {
        get =>
            this.Resolve(
                static alignment => alignment.TextRotation?.Value,
                this._default.TextRotation.Value
            );
        set =>
            this.Modify(
                static (alignment, textRotation) =>
                    alignment with
                    {
                        TextRotation = new TextRotation(textRotation),
                    },
                value
            );
    }

    bool IXLAlignment.WrapText
    {
        get => this.Resolve(static alignment => alignment.WrapText, this._default.WrapText);
        set =>
            this.Modify(
                static (alignment, wrapText) => alignment with { WrapText = wrapText },
                value
            );
    }

    bool IXLAlignment.TopToBottom
    {
        get =>
            this.Resolve(
                static alignment => alignment.TextRotation == TextRotation.VerticalText,
                false
            );
        set =>
            this.Modify(
                static (alignment, topToBottom) =>
                    alignment with
                    {
                        TextRotation = topToBottom ? TextRotation.VerticalText : TextRotation.None,
                    },
                value
            );
    }

    bool IEquatable<IXLAlignment>.Equals(IXLAlignment other) => throw new NotSupportedException();

    IXLStyle IXLAlignment.SetHorizontal(XLAlignmentHorizontalValues value)
    {
        (this as IXLAlignment).Horizontal = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetIndent(int value)
    {
        (this as IXLAlignment).Indent = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetJustifyLastLine() => (this as IXLAlignment).SetJustifyLastLine(true);

    IXLStyle IXLAlignment.SetJustifyLastLine(bool value)
    {
        (this as IXLAlignment).JustifyLastLine = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetReadingOrder(XLAlignmentReadingOrderValues value)
    {
        (this as IXLAlignment).ReadingOrder = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetRelativeIndent(int value)
    {
        (this as IXLAlignment).RelativeIndent = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetShrinkToFit() => (this as IXLAlignment).SetShrinkToFit(true);

    IXLStyle IXLAlignment.SetShrinkToFit(bool value)
    {
        (this as IXLAlignment).ShrinkToFit = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetTextRotation(int value)
    {
        (this as IXLAlignment).TextRotation = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetTopToBottom() => (this as IXLAlignment).SetTopToBottom(true);

    IXLStyle IXLAlignment.SetTopToBottom(bool value)
    {
        (this as IXLAlignment).TopToBottom = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetVertical(XLAlignmentVerticalValues value)
    {
        (this as IXLAlignment).Vertical = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetWrapText() => (this as IXLAlignment).SetWrapText(true);

    IXLStyle IXLAlignment.SetWrapText(bool value)
    {
        (this as IXLAlignment).WrapText = value;
        return this._parent;
    }

    internal void SetValue(IXLAlignment value) =>
        this._parent.ModifyAlignment(
            static (alignment, value) =>
                alignment with
                {
                    Horizontal = value.Horizontal,
                    Vertical = value.Vertical,
                    TextRotation = new TextRotation(value.TextRotation),
                    WrapText = value.WrapText,
                    Indent = value.Indent,
                    RelativeIndent = value.RelativeIndent,
                    JustifyLastLine = value.JustifyLastLine,
                    ShrinkToFit = value.ShrinkToFit,
                    ReadingOrder = value.ReadingOrder,
                },
            value
        );

    private T Resolve<T>(Func<XLDifferentialAlignmentValue, T?> getProperty, T defaultValue)
        where T : struct =>
        this._parent.Resolve(static format => format.Alignment, getProperty) ?? defaultValue;

    private void Modify<T>(
        Func<XLDifferentialAlignmentValue, T, XLDifferentialAlignmentValue> modify,
        T value
    ) => this._parent.ModifyAlignment(modify, value);
}

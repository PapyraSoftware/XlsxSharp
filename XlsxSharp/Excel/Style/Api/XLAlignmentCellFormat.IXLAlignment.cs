using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// Explicit implementation of the <see cref="IXLAlignment"/> interface.
/// </summary>
internal sealed partial class XLAlignmentCellFormat : IXLAlignment
{
    XLAlignmentHorizontalValues IXLAlignment.Horizontal
    {
        get => this.Resolve(static format => format.Alignment.Horizontal);
        set =>
            this.Modify(
                static (alignment, hAlign) => alignment with { Horizontal = hAlign },
                value
            );
    }

    XLAlignmentVerticalValues IXLAlignment.Vertical
    {
        get => this.Resolve(static format => format.Alignment.Vertical);
        set =>
            this.Modify(static (alignment, vAlign) => alignment with { Vertical = vAlign }, value);
    }

    int IXLAlignment.Indent
    {
        get => this.Resolve(static format => format.Alignment.Indent);
        set =>
            this.Modify(
                static (alignment, indent) =>
                {
                    if (alignment.Horizontal == XLAlignmentHorizontalValues.General)
                    {
                        alignment = alignment with
                        {
                            Horizontal = XLAlignmentHorizontalValues.Left,
                        };
                    }

                    if (
                        indent > 0
                        && !(
                            alignment.Horizontal == XLAlignmentHorizontalValues.Left
                            || alignment.Horizontal == XLAlignmentHorizontalValues.Right
                            || alignment.Horizontal == XLAlignmentHorizontalValues.Distributed
                        )
                    )
                    {
                        throw new InvalidOperationException(
                            "For indents, only left, right, and distributed horizontal alignments are supported."
                        );
                    }

                    return alignment with
                    {
                        Indent = indent,
                    };
                },
                value
            );
    }

    bool IXLAlignment.JustifyLastLine
    {
        get => this.Resolve(static format => format.Alignment.JustifyLastLine);
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
        get => this.Resolve(static format => format.Alignment.ReadingOrder);
        set =>
            this.Modify(
                static (alignment, readingOrder) => alignment with { ReadingOrder = readingOrder },
                value
            );
    }

    int IXLAlignment.RelativeIndent
    {
        get => this.Resolve(static format => format.Alignment.RelativeIndent);
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
        get => this.Resolve(static format => format.Alignment.ShrinkToFit);
        set =>
            this.Modify(
                static (alignment, shrinkToFit) => alignment with { ShrinkToFit = shrinkToFit },
                value
            );
    }

    int IXLAlignment.TextRotation
    {
        get => this.Resolve(static format => format.Alignment.TextRotation.Value);
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
        get => this.Resolve(static format => format.Alignment.WrapText);
        set =>
            this.Modify(
                static (alignment, wrapText) => alignment with { WrapText = wrapText },
                value
            );
    }

    bool IXLAlignment.TopToBottom
    {
        get =>
            this.Resolve(static format =>
                format.Alignment.TextRotation == TextRotation.VerticalText
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

    IXLStyle IXLAlignment.SetHorizontal(XLAlignmentHorizontalValues value)
    {
        (this as IXLAlignment).Horizontal = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetVertical(XLAlignmentVerticalValues value)
    {
        (this as IXLAlignment).Vertical = value;
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

    IXLStyle IXLAlignment.SetWrapText() => (this as IXLAlignment).SetWrapText(true);

    IXLStyle IXLAlignment.SetWrapText(bool value)
    {
        (this as IXLAlignment).WrapText = value;
        return this._parent;
    }

    IXLStyle IXLAlignment.SetTopToBottom() => (this as IXLAlignment).SetTopToBottom(true);

    IXLStyle IXLAlignment.SetTopToBottom(bool value)
    {
        (this as IXLAlignment).TopToBottom = value;
        return this._parent;
    }

    bool IEquatable<IXLAlignment>.Equals(IXLAlignment? other)
    {
        if (other == null)
        {
            return false;
        }

        XLAlignmentFormatValue align = this.Resolve(static x => x.Alignment);
        if (align.Horizontal != other.Horizontal)
        {
            return false;
        }

        if (align.Vertical != other.Vertical)
        {
            return false;
        }

        if (align.Indent != other.Indent)
        {
            return false;
        }

        if (align.JustifyLastLine != other.JustifyLastLine)
        {
            return false;
        }

        if (align.ReadingOrder != other.ReadingOrder)
        {
            return false;
        }

        if (align.RelativeIndent != other.RelativeIndent)
        {
            return false;
        }

        if (align.ShrinkToFit != other.ShrinkToFit)
        {
            return false;
        }

        if (align.TextRotation.Value != other.TextRotation)
        {
            return false;
        }

        if (align.WrapText != other.WrapText)
        {
            return false;
        }

        return true;
    }
}

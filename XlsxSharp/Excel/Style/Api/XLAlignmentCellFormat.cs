using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal partial class XLAlignmentCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLAlignmentCellFormat(XLCellFormat parent) => this._parent = parent;

    public override bool Equals(object? obj) =>
        obj is IXLAlignment other && (this as IEquatable<IXLAlignment>).Equals(other);

    public override int GetHashCode() => 0;

    internal void SetValue(IXLAlignment value) =>
        this.Modify(
            static (alignment, other) =>
                alignment with
                {
                    Horizontal = other.Horizontal,
                    Vertical = other.Vertical,
                    TextRotation = new TextRotation(other.TextRotation),
                    WrapText = other.WrapText,
                    Indent = other.Indent,
                    RelativeIndent = other.RelativeIndent,
                    JustifyLastLine = other.JustifyLastLine,
                    ShrinkToFit = other.ShrinkToFit,
                    ReadingOrder = other.ReadingOrder,
                },
            value
        );

    private T Resolve<T>(Func<XLCellFormatValue, T> selector) => this._parent.Resolve(selector);

    private void Modify<TProperty>(
        Func<XLAlignmentFormatValue, TProperty, XLAlignmentFormatValue> modifyAlignment,
        TProperty value
    ) => this._parent.ModifyAlignment(modifyAlignment, value);
}

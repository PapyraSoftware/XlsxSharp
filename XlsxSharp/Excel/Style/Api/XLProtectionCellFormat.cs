using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal sealed partial class XLProtectionCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLProtectionCellFormat(XLCellFormat parent) => this._parent = parent;

    public override bool Equals(object? obj) =>
        obj is IXLProtection other && (this as IEquatable<IXLProtection>).Equals(other);

    public override int GetHashCode() => 0;

    internal void SetValue(IXLProtection value) =>
        this.Modify(
            static (_, other) =>
                new XLProtectionFormatValue { Hidden = other.Hidden, Locked = other.Locked },
            value
        );

    private T Resolve<T>(Func<XLCellFormatValue, T> selector) => this._parent.Resolve(selector);

    private void Modify<TProperty>(
        Func<XLProtectionFormatValue, TProperty, XLProtectionFormatValue> modifyProtection,
        TProperty value
    ) => this._parent.ModifyProtection(modifyProtection, value);
}

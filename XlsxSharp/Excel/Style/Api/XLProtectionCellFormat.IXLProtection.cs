using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal partial class XLProtectionCellFormat : IXLProtection
{
    bool IXLProtection.Locked
    {
        get => this.Resolve(static style => style.Protection.Locked);
        set =>
            this.Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    bool IXLProtection.Hidden
    {
        get => this.Resolve(static style => style.Protection.Hidden);
        set =>
            this.Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXLStyle IXLProtection.SetLocked()
    {
        return (this as IXLProtection).SetLocked(true);
    }

    IXLStyle IXLProtection.SetLocked(bool value)
    {
        (this as IXLProtection).Locked = value;
        return this._parent;
    }

    IXLStyle IXLProtection.SetHidden()
    {
        return (this as IXLProtection).SetHidden(true);
    }

    IXLStyle IXLProtection.SetHidden(bool value)
    {
        (this as IXLProtection).Hidden = value;
        return this._parent;
    }

    bool IEquatable<IXLProtection>.Equals(IXLProtection? other)
    {
        if (other is null)
        {
            return false;
        }

        XLProtectionFormatValue protection = this.Resolve(static style => style.Protection);
        return protection.Locked == other.Locked && protection.Hidden == other.Hidden;
    }
}

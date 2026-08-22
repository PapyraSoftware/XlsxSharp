using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// An API object to modify protection of a dxf.
/// </summary>
internal class XLDxfProtectionFormat : IXLProtection
{
    private readonly XLDxFormat _parent;

    public XLDxfProtectionFormat(XLDxFormat parent)
    {
        this._parent = parent;
    }

    bool IXLProtection.Locked
    {
        get => this.Locked;
        set => this.Locked = value;
    }

    bool IXLProtection.Hidden
    {
        get => this.Hidden;
        set => this.Hidden = value;
    }

    private bool Locked
    {
        get => this.Resolve(p => p.Locked, true);
        set =>
            this.Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    private bool Hidden
    {
        get => this.Resolve(p => p.Hidden, false);
        set =>
            this.Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXLStyle IXLProtection.SetLocked()
    {
        this.Locked = true;
        return this._parent;
    }

    IXLStyle IXLProtection.SetLocked(bool value)
    {
        this.Locked = value;
        return this._parent;
    }

    IXLStyle IXLProtection.SetHidden()
    {
        this.Hidden = true;
        return this._parent;
    }

    IXLStyle IXLProtection.SetHidden(bool value)
    {
        this.Hidden = value;
        return this._parent;
    }

    bool IEquatable<IXLProtection>.Equals(IXLProtection other)
    {
        return this.Locked == other.Locked && this.Hidden == other.Hidden;
    }

    internal void SetValue(IXLProtection value)
    {
        this.Modify(
            (_, other) =>
                new XLDifferentialProtectionValue { Locked = other.Locked, Hidden = other.Hidden },
            value
        );
    }

    private T Resolve<T>(Func<XLDifferentialProtectionValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return this._parent.Resolve(static format => format.Protection, getProperty)
            ?? defaultValue;
    }

    private void Modify<T>(
        Func<XLDifferentialProtectionValue, T, XLDifferentialProtectionValue> modify,
        T value
    )
    {
        this._parent.ModifyProtection(modify, value);
    }
}

using System;
using ClosedXML.Excel.Formatting;

namespace ClosedXML.Excel;

/// <summary>
/// An API object to modify protection of a dxf.
/// </summary>
internal class XLDxfProtectionFormat : IXLProtection
{
    private readonly XLDxFormat _parent;

    public XLDxfProtectionFormat(XLDxFormat parent)
    {
        _parent = parent;
    }

    bool IXLProtection.Locked
    {
        get => Locked;
        set => Locked = value;
    }

    bool IXLProtection.Hidden
    {
        get => Hidden;
        set => Hidden = value;
    }

    private bool Locked
    {
        get => Resolve(p => p.Locked, true);
        set => Modify(static (protection, locked) => protection with { Locked = locked }, value);
    }

    private bool Hidden
    {
        get => Resolve(p => p.Hidden, false);
        set => Modify(static (protection, hidden) => protection with { Hidden = hidden }, value);
    }

    IXLStyle IXLProtection.SetLocked()
    {
        Locked = true;
        return _parent;
    }

    IXLStyle IXLProtection.SetLocked(bool value)
    {
        Locked = value;
        return _parent;
    }

    IXLStyle IXLProtection.SetHidden()
    {
        Hidden = true;
        return _parent;
    }

    IXLStyle IXLProtection.SetHidden(bool value)
    {
        Hidden = value;
        return _parent;
    }

    bool IEquatable<IXLProtection>.Equals(IXLProtection other)
    {
        return Locked == other.Locked && Hidden == other.Hidden;
    }

    internal void SetValue(IXLProtection value)
    {
        Modify((_, other) => new XLDifferentialProtectionValue
        {
            Locked = other.Locked,
            Hidden = other.Hidden
        }, value);
    }

    private T Resolve<T>(Func<XLDifferentialProtectionValue, T?> getProperty, T defaultValue)
        where T : struct
    {
        return _parent.Resolve(static format => format.Protection, getProperty) ?? defaultValue;
    }

    private void Modify<T>(Func<XLDifferentialProtectionValue, T, XLDifferentialProtectionValue> modify, T value)
    {
        _parent.ModifyProtection(modify, value);
    }
}

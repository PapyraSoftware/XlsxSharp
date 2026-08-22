using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

public class FormatTestCase<TApi>
{
    private readonly Func<TApi, object> _getter;
    private readonly Action<TApi, object> _setter;
    private readonly IReadOnlyList<object> _testValues;

    private FormatTestCase(
        Func<TApi, object> getter,
        Action<TApi, object> setter,
        params object[] testValues
    )
    {
        this._getter = getter;
        this._setter = setter;
        this._testValues = testValues;
    }

    internal static FormatTestCase<IXLFont> ForFont<T>(
        Func<IXLFont, T> getter,
        Action<IXLFont, T> setter,
        params T[] testValues
    )
    {
        return new FormatTestCase<IXLFont>(
            font => getter(font),
            (font, value) => setter(font, (T)value),
            [.. testValues.Cast<object>()]
        );
    }

    internal static FormatTestCase<IXLFill> ForFill<T>(
        Func<IXLFill, T> getter,
        Action<IXLFill, T> setter,
        params T[] testValues
    )
    {
        return new FormatTestCase<IXLFill>(
            fill => getter(fill),
            (fill, value) => setter(fill, (T)value),
            [.. testValues.Cast<object>()]
        );
    }

    internal static FormatTestCase<IXLBorder> ForBorder<T>(
        Func<IXLBorder, T> getter,
        Action<IXLBorder, T> setter,
        params T[] testValues
    )
    {
        return new FormatTestCase<IXLBorder>(
            border => getter(border),
            (border, value) => setter(border, (T)value),
            [.. testValues.Cast<object>()]
        );
    }

    internal static FormatTestCase<IXLAlignment> ForAlignment<T>(
        Func<IXLAlignment, T> getter,
        Action<IXLAlignment, T> setter,
        params T[] testValues
    )
    {
        return new FormatTestCase<IXLAlignment>(
            align => getter(align),
            (align, value) => setter(align, (T)value),
            [.. testValues.Cast<object>()]
        );
    }

    internal static FormatTestCase<IXLProtection> ForProtection<T>(
        Func<IXLProtection, T> getter,
        Action<IXLProtection, T> setter,
        params T[] testValues
    )
    {
        return new FormatTestCase<IXLProtection>(
            protection => getter(protection),
            (protection, value) => setter(protection, (T)value),
            [.. testValues.Cast<object>()]
        );
    }

    internal IEnumerable<object> Values => this._testValues;

    internal object GetPropertyValue(TApi font)
    {
        return this._getter(font);
    }

    internal void SetPropertyValue(TApi font, object value)
    {
        this._setter(font, value);
    }
}

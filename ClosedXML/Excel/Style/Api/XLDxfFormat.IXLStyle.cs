using System;

namespace ClosedXML.Excel;

internal partial class XLDxFormat : IXLStyle
{
    IXLAlignment IXLStyle.Alignment
    {
        get => Alignment;
        set => Alignment.SetValue(value);
    }

    IXLBorder IXLStyle.Border
    {
        get => Border;
        set => Border.SetValue(value);
    }

    IXLNumberFormat IXLStyle.DateFormat => throw new NotImplementedException();

    IXLFill IXLStyle.Fill
    {
        get => Fill;
        set => Fill.SetValue(value);
    }

    IXLFont IXLStyle.Font
    {
        get => Font;
        set => Font.SetValue(value);
    }

    bool IXLStyle.IncludeQuotePrefix
    {
        get => false;
        set => throw new NotSupportedException($"Differential format doesn't support {nameof(IXLStyle.IncludeQuotePrefix)}.");
    }

    IXLNumberFormat IXLStyle.NumberFormat
    {
        get => NumberFormat;
        set => NumberFormat.SetValue(value);
    }

    IXLProtection IXLStyle.Protection
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    IXLStyle IXLStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        (this as IXLStyle).IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    bool IEquatable<IXLStyle>.Equals(IXLStyle? other)
    {
        throw new NotSupportedException();
    }
}

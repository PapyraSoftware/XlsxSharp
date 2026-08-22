using System;

namespace XlsxSharp.Excel;

internal partial class XLDxFormat : IXLStyle
{
    IXLAlignment IXLStyle.Alignment
    {
        get => this.Alignment;
        set => this.Alignment.SetValue(value);
    }

    IXLBorder IXLStyle.Border
    {
        get => this.Border;
        set => this.Border.SetValue(value);
    }

    IXLNumberFormat IXLStyle.DateFormat => this.NumberFormat;

    IXLFill IXLStyle.Fill
    {
        get => this.Fill;
        set => this.Fill.SetValue(value);
    }

    IXLFont IXLStyle.Font
    {
        get => this.Font;
        set => this.Font.SetValue(value);
    }

    bool IXLStyle.IncludeQuotePrefix
    {
        get => false;
        set =>
            throw new NotSupportedException(
                $"Differential format doesn't support {nameof(IXLStyle.IncludeQuotePrefix)}."
            );
    }

    IXLNumberFormat IXLStyle.NumberFormat
    {
        get => this.NumberFormat;
        set => this.NumberFormat.SetValue(value);
    }

    IXLProtection IXLStyle.Protection
    {
        get => this.Protection;
        set => this.Protection.SetValue(value);
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

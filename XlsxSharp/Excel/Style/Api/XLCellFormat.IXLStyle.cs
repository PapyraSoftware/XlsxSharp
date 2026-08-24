using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// Methods and properties of <see cref="IXLStyle"/>. The <see cref="XLCellFormat"/> has many
/// properties with same name, but different type, so the interface is explicitly implemented.
/// </summary>
internal partial class XLCellFormat : IXLStyle
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
        set => this.Font.SetFont(value);
    }

    bool IXLStyle.IncludeQuotePrefix
    {
        get => this.IncludeQuotePrefix;
        set => this.IncludeQuotePrefix = value;
    }

    IXLNumberFormat IXLStyle.NumberFormat
    {
        get => this.NumberFormat;
        set => this.NumberFormat.SetNumberFormat(value.Format);
    }

    IXLProtection IXLStyle.Protection
    {
        get => this.Protection;
        set => this.Protection.SetValue(value);
    }

    IXLStyle IXLStyle.SetIncludeQuotePrefix(bool includeQuotePrefix)
    {
        this.IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    bool IEquatable<IXLStyle>.Equals(IXLStyle? other)
    {
        if (other is null)
        {
            return false;
        }

        // The API object for each component implement IEquitable<IXL..>
        if (!this.NumberFormat.Equals(other.NumberFormat))
        {
            return false;
        }

        if (!this.Font.Equals(other.Font))
        {
            return false;
        }

        if (!this.IncludeQuotePrefix.Equals(other.IncludeQuotePrefix))
        {
            return false;
        }

        if (!this.Fill.Equals(other.Fill))
        {
            return false;
        }

        if (!this.Border.Equals(other.Border))
        {
            return false;
        }

        if (!this.Alignment.Equals(other.Alignment))
        {
            return false;
        }

        if (!this.Protection.Equals(other.Protection))
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// A helper method that is used when a style if copied from one object to another.
    /// For example, <c>rangeApi.Style = someOtherApi.Style</c>.
    /// </summary>
    internal void SetStyle(IXLStyle value)
    {
        if (value is not XLCellFormat cellFormat)
        {
            throw new NotSupportedException("Can only copy cell format style.");
        }

        XLCellFormatValue otherCellFormat = cellFormat._formatValue.Resolve();
        XLCellFormatValue registeredCellFormat = this._workbook.Styles.GetRegisteredCellFormat(
            otherCellFormat
        );
        this.ModifyFormat((_, cellForamt) => cellForamt, registeredCellFormat);
    }
}

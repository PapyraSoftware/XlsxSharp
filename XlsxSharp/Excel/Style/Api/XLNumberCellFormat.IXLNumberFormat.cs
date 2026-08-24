using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal sealed partial class XLNumberCellFormat : IXLNumberFormat
{
    int IXLNumberFormatBase.NumberFormatId
    {
        get => this.NumberFormatId;
        set => this.NumberFormatId = value;
    }

    string IXLNumberFormatBase.Format
    {
        get => this.Format;
        set => this.Format = XLNumberFormat.Parse(value);
    }

    bool IEquatable<IXLNumberFormatBase>.Equals(IXLNumberFormatBase? other)
    {
        if (other is null)
        {
            return false;
        }

        return other.Format == this.Format;
    }

    IXLStyle IXLNumberFormat.SetNumberFormatId(int value)
    {
        this.NumberFormatId = value;
        return this._parent;
    }

    IXLStyle IXLNumberFormat.SetFormat(string value)
    {
        this.Format = XLNumberFormat.Parse(value);
        return this._parent;
    }
}

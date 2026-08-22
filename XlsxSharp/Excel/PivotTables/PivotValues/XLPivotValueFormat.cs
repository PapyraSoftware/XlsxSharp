using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.PivotValues;

internal class XLPivotValueFormat : IXLPivotValueFormat
{
    private readonly XLPivotDataField _pivotValue;

    public XLPivotValueFormat(XLPivotDataField pivotValue)
    {
        this._pivotValue = pivotValue;
    }

    public Int32 NumberFormatId
    {
        get
        {
            if (this._pivotValue.NumberFormatValue is null)
            {
                return -1;
            }

            if (
                !XLPredefinedFormat.NumberFormatIds.TryGetValue(
                    this._pivotValue.NumberFormatValue.Value,
                    out int numFmtId
                )
            )
            {
                return -1;
            }

            return numFmtId;
        }
        set
        {
            if (!XLPredefinedFormat.FormatCodes.TryGetValue(value, out XLNumberFormat format))
            {
                throw new ArgumentOutOfRangeException(
                    $"Only predefined format is permitted. Use nested enums/members of {nameof(XLPredefinedFormat)}."
                );
            }

            this._pivotValue.NumberFormatValue = format;
        }
    }

    public String Format
    {
        get => this._pivotValue.NumberFormatValue ?? string.Empty;
        set { this._pivotValue.NumberFormatValue = XLNumberFormat.Parse(value); }
    }

    public IXLPivotValue SetNumberFormatId(Int32 value)
    {
        this.NumberFormatId = value;
        return this._pivotValue;
    }

    public IXLPivotValue SetFormat(String value)
    {
        this.Format = value;
        return this._pivotValue;
    }
}

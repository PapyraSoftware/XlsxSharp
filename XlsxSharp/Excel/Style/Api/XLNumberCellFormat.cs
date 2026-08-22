using System;
using System.Collections.Generic;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// An API object to modify number format of a <see cref="XLCellFormat">cell format</see>.
/// </summary>
internal sealed partial class XLNumberCellFormat
{
    private readonly XLCellFormat _parent;

    internal XLNumberCellFormat(XLCellFormat parent) => this._parent = parent;

    private int NumberFormatId
    {
        get
        {
            XLNumberFormat numberFormat = this._parent.Resolve(static x => x.NumberFormat);
            return XLPredefinedFormat.NumberFormatIds.GetValueOrDefault(numberFormat, -1);
        }
        set
        {
            if (!XLPredefinedFormat.FormatCodes.TryGetValue(value, out XLNumberFormat format))
            {
                throw new ArgumentOutOfRangeException(
                    $"Only predefined format is permitted. Use nested enums/members of {nameof(XLPredefinedFormat)}."
                );
            }

            this.Format = format;
        }
    }

    private XLNumberFormat Format
    {
        get => this._parent.Resolve(static x => x.NumberFormat);
        set => this._parent.ModifyNumberFormat(value);
    }

    public override bool Equals(object? obj) =>
        obj is IXLNumberFormatBase other && (this as IEquatable<IXLNumberFormatBase>).Equals(other);

    public override int GetHashCode() => 0;

    internal void SetNumberFormat(string numberFormat) =>
        this.Format = XLNumberFormat.Parse(numberFormat);
}

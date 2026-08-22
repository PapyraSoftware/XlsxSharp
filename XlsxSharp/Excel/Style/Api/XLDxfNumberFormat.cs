using System;
using System.Collections.Generic;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// An API object to modify number format of a dxf.
/// </summary>
internal class XLDxfNumberFormat : IXLNumberFormat
{
    private readonly XLDxFormat _parent;

    internal XLDxfNumberFormat(XLDxFormat parent) => this._parent = parent;

    /// <inheritdoc />
    int IXLNumberFormatBase.NumberFormatId
    {
        get => this.NumberFormatId;
        set => this.NumberFormatId = value;
    }

    /// <inheritdoc />
    string IXLNumberFormatBase.Format
    {
        get => this.Format;
        set => this.Format = value;
    }

    private int NumberFormatId
    {
        get
        {
            string? numberFormat = this._parent.Resolve(static x => x.NumberFormat, static x => x);
            if (numberFormat is null)
            {
                return XLPredefinedFormat.General;
            }

            return XLPredefinedFormat.NumberFormatIds.GetValueOrDefault(
                new XLNumberFormat(numberFormat),
                -1
            );
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

    private string Format
    {
        get
        {
            string? format = this._parent.Resolve(static x => x.NumberFormat, static x => x);
            if (format is null)
            {
                return XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General];
            }

            return format;
        }
        set => this._parent.ModifyNumberFormat(XLNumberFormat.Parse(value));
    }

    IXLStyle IXLNumberFormat.SetNumberFormatId(int value)
    {
        this.NumberFormatId = value;
        return this._parent;
    }

    IXLStyle IXLNumberFormat.SetFormat(string value)
    {
        this.Format = value;
        return this._parent;
    }

    bool IEquatable<IXLNumberFormatBase>.Equals(IXLNumberFormatBase other) =>
        this.Format == other.Format;

    internal void SetValue(IXLNumberFormat value) => this.Format = value.Format;
}

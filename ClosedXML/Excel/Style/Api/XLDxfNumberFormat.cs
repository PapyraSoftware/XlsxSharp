using System;
using System.Collections.Generic;

namespace ClosedXML.Excel;

/// <summary>
/// An API object to modify number format of a dxf.
/// </summary>
internal class XLDxfNumberFormat : IXLNumberFormat
{
    private readonly XLDxFormat _parent;

    internal XLDxfNumberFormat(XLDxFormat parent)
    {
        _parent = parent;
    }

    /// <inheritdoc />
    int IXLNumberFormatBase.NumberFormatId
    {
        get => NumberFormatId;
        set => NumberFormatId = value;
    }

    /// <inheritdoc />
    string IXLNumberFormatBase.Format
    {
        get => Format;
        set => Format = value;
    }

    private int NumberFormatId
    {
        get
        {
            var numberFormat = _parent.Resolve(static x => x.NumberFormat, static x => x);
            if (numberFormat is null)
                return XLPredefinedFormat.General;

            return XLPredefinedFormat.NumberFormatIds.GetValueOrDefault(new XLNumberFormat(numberFormat), -1);
        }
        set
        {
            if (!XLPredefinedFormat.FormatCodes.TryGetValue(value, out var format))
                throw new ArgumentOutOfRangeException($"Only predefined format is permitted. Use nested enums/members of {nameof(XLPredefinedFormat)}.");

            Format = format;
        }
    }

    private string Format
    {
        get
        {
            var format = _parent.Resolve(static x => x.NumberFormat, static x => x);
            if (format is null)
                return XLPredefinedFormat.FormatCodes[XLPredefinedFormat.General];

            return format;
        }
        set => _parent.ModifyNumberFormat(XLNumberFormat.Parse(value));
    }

    IXLStyle IXLNumberFormat.SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _parent;
    }

    IXLStyle IXLNumberFormat.SetFormat(string value)
    {
        Format = value;
        return _parent;
    }

    bool IEquatable<IXLNumberFormatBase>.Equals(IXLNumberFormatBase other)
    {
        return Format == other.Format;
    }

    internal void SetValue(IXLNumberFormat value)
    {
        Format = value.Format;
    }
}

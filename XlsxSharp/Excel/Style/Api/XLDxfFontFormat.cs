using System;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal partial class XLDxfFontFormat
{
    private readonly XLDxFormat _parent;

    internal XLDxfFontFormat(XLDxFormat parent) => this._parent = parent;

    internal void SetValue(IXLFont value) =>
        this._parent.ModifyFont(
            static (font, value) =>
                font with
                {
                    Bold = value.Bold,
                    Italic = value.Italic,
                    Underline = value.Underline,
                    Strikethrough = value.Strikethrough,
                    VerticalAlignment = value.VerticalAlignment,
                    Shadow = value.Shadow,
                    Size = XLFontSize.FromPoints(value.FontSize),
                    Color = value.FontColor,
                    Name = value.FontName,
                    Family = value.FontFamilyNumbering,
                    Charset = value.FontCharSet,
                    Scheme = value.FontScheme,
                },
            value
        );

    private T Resolve<T>(Func<XLDifferentialFontValue, T?> getProperty, T defaultValue)
        where T : struct =>
        this._parent.Resolve(static format => format.Font, getProperty) ?? defaultValue;

    private T Resolve<T>(Func<XLDifferentialFontValue, T?> getProperty, T defaultValue)
        where T : class =>
        this._parent.Resolve(static format => format.Font, getProperty) ?? defaultValue;

    private void Modify<T>(
        Func<XLDifferentialFontValue, T, XLDifferentialFontValue> modify,
        T value
    ) => this._parent.ModifyFont(modify, value);
}

using System.Globalization;
using XlsxSharp.ExcelNumberFormat;

namespace XlsxSharp.Extensions;

internal static class FormatExtensions
{
    public static string ToExcelFormat(this object o, string format, CultureInfo culture)
    {
        NumberFormat nf = new(format);
        if (!nf.IsValid)
        {
            return format;
        }

        return nf.Format(o, culture);
    }
}

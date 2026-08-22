using System;
using static XlsxSharp.Excel.XLPredefinedFormat.DateTime;

namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// A strongly typed number format.
/// </summary>
internal readonly record struct XLNumberFormat : IEquatable<string>
{
    private readonly string _format;

    internal XLNumberFormat(string format) =>
        this._format = format ?? throw new ArgumentNullException(nameof(format));

    public string Format => this._format;

    public static implicit operator string(XLNumberFormat value) => value._format;

    /// <summary>
    /// Is a number format a general number format?
    /// </summary>
    internal bool IsGeneralFormat() =>
        // General format is an empty string.
        this._format.Length == 0;

    internal static XLNumberFormat Parse(string formatCode) =>
        // Format code was originally just a string and not checked, so keep the same semantic for now.
        new(formatCode);

    public bool Equals(string other) => other == this._format;

    internal XLDataType GetNumberDataType()
    {
        if (XLPredefinedFormat.NumberFormatIds.TryGetValue(this, out int formatId))
        {
            if (IsTimeOnlyFormat(formatId))
            {
                return XLDataType.TimeSpan;
            }

            if (IsDateTimeFormat(formatId))
            {
                return XLDataType.DateTime;
            }

            return XLDataType.Number;
        }

        if (!string.IsNullOrWhiteSpace(this._format))
        {
            XLDataType? dataType = this.GetDataTypeFromFormat();
            return dataType ?? XLDataType.Number;
        }

        return XLDataType.Number;
    }

    private static bool IsDateTimeFormat(int numberFormatId) =>
        (XLPredefinedFormat.DateTime)numberFormatId
            is DayMonthYear4WithSlashes
                or DayMonthAbbrYear2WithDashes
                or DayMonthAbbrWithDash
                or MonthDayYear4WithDashesHour24Minutes;

    private static bool IsTimeOnlyFormat(int numberFormatId) =>
        (XLPredefinedFormat.DateTime)numberFormatId
            is Hour12MinutesAmPm
                or Hour12MinutesSecondsAmPm
                or Hour24Minutes
                or Hour24MinutesSeconds
                or MinutesSeconds
                or Hour12MinutesSeconds
                or MinutesSecondsMillis1;

    private XLDataType? GetDataTypeFromFormat()
    {
        int length = this._format.Length;
        string f = this._format.ToLowerInvariant();
        for (int i = 0; i < length; i++)
        {
            char c = f[i];
            if (c == '"')
            {
                i = f.IndexOf('"', i + 1);
            }
            else if (c == '[')
            {
                // #1742 We need to skip locale prefixes in DateTime formats [...]
                i = f.IndexOf(']', i + 1);
                if (i == -1)
                {
                    return null;
                }
            }
            else if (c is '0' or '#' or '?')
            {
                return XLDataType.Number;
            }
            else if (c is 'y' or 'd')
            {
                return XLDataType.DateTime;
            }
            else if (c is 'h' or 's')
            {
                return XLDataType.TimeSpan;
            }
            else if (c == 'm')
            {
                // Excel treats "m" immediately after "hh" or "h" or immediately before "ss" or "s" as minutes, otherwise as a month value
                // We can ignore the "hh" or "h" prefixes as these would have been detected by the preceding condition above.
                // So we just need to make sure any 'm' is followed immediately by "ss" or "s" (excluding placeholders) to detect a timespan value
                for (int j = i + 1; j < length; j++)
                {
                    switch (f[j])
                    {
                        case 'm':
                            continue;
                        case 's':
                            return XLDataType.TimeSpan;
                        case >= 'a' and <= 'z':
                        case >= '0' and <= '9':
                            return XLDataType.DateTime;
                    }
                }
                return XLDataType.DateTime;
            }
        }
        return null;
    }
}

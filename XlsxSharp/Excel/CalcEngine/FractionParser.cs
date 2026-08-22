#nullable disable

using System.Globalization;
using System.Text.RegularExpressions;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// Parse a fraction for text-to-number type coercion.
/// </summary>
internal static class FractionParser
{
    private static readonly Regex FractionRegex = new(
        @"^ *([+-]?) *([0-9]+) ([0-9]{1,5})/([0-9]{1,5}) *$",
        RegexOptions.CultureInvariant
    );

    public static bool TryParse(string s, out double result)
    {
        result = default;
        Match match = FractionRegex.Match(s);
        if (!match.Success)
        {
            return false;
        }

        int denominator = ParseInt(match.Groups[4]);
        if (denominator == 0 || denominator > short.MaxValue)
        {
            return false;
        }

        int numerator = ParseInt(match.Groups[3]);
        if (numerator > short.MaxValue)
        {
            return false;
        }

        Group sign = match.Groups[1];
        int wholeNumber = ParseInt(match.Groups[2]);

        double fraction = wholeNumber + numerator / (double)denominator;
        bool hasNegativeSign = sign.Success && sign.Value.Length > 0 && sign.Value[0] == '-';
        result = hasNegativeSign ? -fraction : fraction;
        return true;

        static int ParseInt(Capture capture) =>
            int.Parse(capture.Value, NumberStyles.None, CultureInfo.InvariantCulture);
    }
}

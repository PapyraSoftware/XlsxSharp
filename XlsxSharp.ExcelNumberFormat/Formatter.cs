using System.Globalization;
using System.Text;

namespace XlsxSharp.ExcelNumberFormat;

internal static class Formatter
{
    public static string Format(
        object value,
        string formatString,
        CultureInfo culture,
        bool isDate1904
    )
    {
        NumberFormat format = new(formatString);
        if (!format.IsValid)
        {
            return CompatibleConvert.ToString(value, culture);
        }

        Section? section = Evaluator.GetSection(format.Sections, value);
        if (section == null)
        {
            return CompatibleConvert.ToString(value, culture);
        }

        return Format(value, section, culture, isDate1904);
    }

    public static string Format(object value, Section node, CultureInfo culture, bool isDate1904)
    {
        switch (node.Type)
        {
            case SectionType.Number:
                // Hide sign under certain conditions and section index
                double number = Convert.ToDouble(value, culture);
                if ((node.SectionIndex == 0 && node.Condition != null) || node.SectionIndex == 1)
                {
                    number = Math.Abs(number);
                }

                return FormatNumber(number, node.Number!, culture);

            case SectionType.Date:
                if (
                    ExcelDateTime.TryConvert(
                        value,
                        isDate1904,
                        culture,
                        out ExcelDateTime excelDateTime
                    )
                )
                {
                    return FormatDate(excelDateTime, node.GeneralTextDateDurationParts!, culture);
                }
                else
                {
                    throw new FormatException("Unexpected date value");
                }

            case SectionType.Duration:
                if (value is TimeSpan ts)
                {
                    return FormatTimeSpan(ts, node.GeneralTextDateDurationParts!, culture);
                }
                else
                {
                    double d = Convert.ToDouble(value);
                    return FormatTimeSpan(
                        TimeSpan.FromDays(d),
                        node.GeneralTextDateDurationParts!,
                        culture
                    );
                }

            case SectionType.General:
            case SectionType.Text:
                return FormatGeneralText(
                    CompatibleConvert.ToString(value, culture),
                    node.GeneralTextDateDurationParts!
                );

            case SectionType.Exponential:
                return FormatExponential(Convert.ToDouble(value, culture), node, culture);

            case SectionType.Fraction:
                return FormatFraction(Convert.ToDouble(value, culture), node, culture);

            default:
                throw new InvalidOperationException("Unknown number format section");
        }
    }

    private static string FormatGeneralText(string text, ReadOnlyMemory<char>[] tokens)
    {
        StringBuilder result = new();
        for (int i = 0; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> token = tokens[i].Span;
            if (Token.IsGeneral(token) || token is "@")
            {
                result.Append(text);
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
        return result.ToString();
    }

    private static string FormatTimeSpan(
        TimeSpan timeSpan,
        ReadOnlyMemory<char>[] tokens,
        CultureInfo culture
    )
    {
        // NOTE/TODO: assumes there is exactly one [hh], [mm] or [ss] using the integer part of TimeSpan.TotalXXX when formatting.
        // The timeSpan input is then truncated to the remainder fraction, which is used to format mm and/or ss.
        StringBuilder result = new();
        bool containsMilliseconds = false;
        for (int i = tokens.Length - 1; i >= 0; i--)
        {
            if (tokens[i].Span.StartsWith(".0"))
            {
                containsMilliseconds = true;
                break;
            }
        }

        for (int i = 0; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> token = tokens[i].Span;

            if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
            {
                int value = timeSpan.Minutes;
                AppendPadded(result, value, token.Length);
            }
            else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
            {
                // If format does not include ms, then include ms in seconds and round before printing
                double formatMs = containsMilliseconds ? 0 : timeSpan.Milliseconds / 1000D;
                int value = (int)
                    Math.Round(timeSpan.Seconds + formatMs, 0, MidpointRounding.AwayFromZero);
                AppendPadded(result, value, token.Length);
            }
            else if (token.StartsWith("[h", StringComparison.OrdinalIgnoreCase))
            {
                int value = (int)timeSpan.TotalHours;
                AppendPadded(result, value, token.Length - 2);
                timeSpan = new TimeSpan(
                    0,
                    0,
                    Math.Abs(timeSpan.Minutes),
                    Math.Abs(timeSpan.Seconds),
                    Math.Abs(timeSpan.Milliseconds)
                );
            }
            else if (token.StartsWith("[m", StringComparison.OrdinalIgnoreCase))
            {
                int value = (int)timeSpan.TotalMinutes;
                AppendPadded(result, value, token.Length - 2);
                timeSpan = new TimeSpan(
                    0,
                    0,
                    0,
                    Math.Abs(timeSpan.Seconds),
                    Math.Abs(timeSpan.Milliseconds)
                );
            }
            else if (token.StartsWith("[s", StringComparison.OrdinalIgnoreCase))
            {
                int value = (int)timeSpan.TotalSeconds;
                AppendPadded(result, value, token.Length - 2);
                timeSpan = new TimeSpan(0, 0, 0, 0, Math.Abs(timeSpan.Milliseconds));
            }
            else if (token.StartsWith(".0"))
            {
                int value = timeSpan.Milliseconds;
                result.Append('.');
                AppendPadded(result, value, token.Length - 1);
            }
            else
            {
                FormatLiteral(token, result);
            }
        }

        return result.ToString();
    }

    private static string FormatDate(
        ExcelDateTime date,
        ReadOnlyMemory<char>[] tokens,
        CultureInfo culture
    )
    {
        bool containsAmPm = ContainsAmPm(tokens);

        StringBuilder result = new();
        for (int i = 0; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> token = tokens[i].Span;

            if (token.StartsWith("y", StringComparison.OrdinalIgnoreCase))
            {
                // year
                int digits = token.Length;
                if (digits < 2)
                {
                    digits = 2;
                }

                if (digits == 3)
                {
                    digits = 4;
                }

                int year = date.Year;
                if (digits == 2)
                {
                    year = year % 100;
                }

                AppendPadded(result, year, digits);
            }
            else if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
            {
                // If  "m" or "mm" code is used immediately after the "h" or "hh" code (for hours) or immediately before
                // the "ss" code (for seconds), the application shall display minutes instead of the month.
                if (LookBackDatePart(tokens, i - 1, "h") || LookAheadDatePart(tokens, i + 1, "s"))
                {
                    AppendPadded(result, date.Minute, token.Length);
                }
                else
                {
                    int digits = token.Length;
                    if (digits == 3)
                    {
                        result.Append(culture.DateTimeFormat.AbbreviatedMonthNames[date.Month - 1]);
                    }
                    else if (digits == 4)
                    {
                        result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1]);
                    }
                    else if (digits == 5)
                    {
                        result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1][0]);
                    }
                    else
                    {
                        AppendPadded(result, date.Month, digits);
                    }
                }
            }
            else if (token.StartsWith("d", StringComparison.OrdinalIgnoreCase))
            {
                int digits = token.Length;
                if (digits == 3)
                {
                    // Sun-Sat
                    result.Append(culture.DateTimeFormat.AbbreviatedDayNames[(int)date.DayOfWeek]);
                }
                else if (digits == 4)
                {
                    // Sunday-Saturday
                    result.Append(culture.DateTimeFormat.DayNames[(int)date.DayOfWeek]);
                }
                else
                {
                    AppendPadded(result, date.Day, digits);
                }
            }
            else if (token.StartsWith("h", StringComparison.OrdinalIgnoreCase))
            {
                int digits = token.Length;
                if (containsAmPm)
                {
                    AppendPadded(result, (date.Hour + 11) % 12 + 1, digits);
                }
                else
                {
                    AppendPadded(result, date.Hour, digits);
                }
            }
            else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
            {
                AppendPadded(result, date.Second, token.Length);
            }
            else if (token.StartsWith("g", StringComparison.OrdinalIgnoreCase))
            {
                int era = culture.DateTimeFormat.Calendar.GetEra(date.AdjustedDateTime);
                int digits = token.Length;
                if (digits < 3)
                {
                    result.Append(culture.DateTimeFormat.GetAbbreviatedEraName(era));
                }
                else
                {
                    result.Append(culture.DateTimeFormat.GetEraName(era));
                }
            }
            else if (token.Equals("am/pm", StringComparison.OrdinalIgnoreCase))
            {
                string ampm = date.ToString("tt", CultureInfo.InvariantCulture);
                result.Append(ampm.ToUpperInvariant());
            }
            else if (token.Equals("a/p", StringComparison.OrdinalIgnoreCase))
            {
                string ampm = date.ToString("%t", CultureInfo.InvariantCulture);
                if (char.IsUpper(token[0]))
                {
                    result.Append(ampm.ToUpperInvariant());
                }
                else
                {
                    result.Append(ampm.ToLowerInvariant());
                }
            }
            else if (token.StartsWith(".0"))
            {
                int value = date.Millisecond;
                result.Append('.');
                AppendPadded(result, value, token.Length - 1);
            }
            else if (token is "/")
            {
#if NETSTANDARD1_0
                result.Append(DateTime.MaxValue.ToString("/d", culture)[0]);
#else
                result.Append(culture.DateTimeFormat.DateSeparator);
#endif
            }
            else if (token is ",")
            {
                while (i < tokens.Length - 1 && tokens[i + 1].Span is ",")
                {
                    i++;
                }

                result.Append(",");
            }
            else
            {
                FormatLiteral(token, result);
            }
        }

        return result.ToString();
    }

    private static bool LookAheadDatePart(
        ReadOnlyMemory<char>[] tokens,
        int fromIndex,
        string startsWith
    )
    {
        for (int i = fromIndex; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> token = tokens[i].Span;
            if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (Token.IsDatePart(token))
            {
                return false;
            }
        }

        return false;
    }

    private static bool LookBackDatePart(
        ReadOnlyMemory<char>[] tokens,
        int fromIndex,
        string startsWith
    )
    {
        for (int i = fromIndex; i >= 0; i--)
        {
            ReadOnlySpan<char> token = tokens[i].Span;
            if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (Token.IsDatePart(token))
            {
                return false;
            }
        }

        return false;
    }

    private static bool ContainsAmPm(ReadOnlyMemory<char>[] tokens)
    {
        foreach (ReadOnlyMemory<char> token in tokens)
        {
            if (token.Span.Equals("am/pm", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (token.Span.Equals("a/p", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Appends <paramref name="value"/> the way <c>value.ToString("D" + digits)</c> would (sign,
    /// then the magnitude zero-padded to at least <paramref name="digits"/> characters), without
    /// allocating the composite format string or an intermediate result string.
    /// </summary>
    private static void AppendPadded(StringBuilder result, int value, int digits)
    {
        Span<char> buffer = stackalloc char[16];
        value.TryFormat(buffer, out int written);
        ReadOnlySpan<char> formatted = buffer[..written];

        // Some cultures use a non-ASCII NegativeSign (e.g. U+2212 for sv-SE), so the sign must be
        // detected via NumberFormatInfo rather than a literal '-'.
        string negativeSign = NumberFormatInfo.CurrentInfo.NegativeSign;
        if (formatted.StartsWith(negativeSign))
        {
            result.Append(negativeSign);
            formatted = formatted[negativeSign.Length..];
        }

        int padding = digits - formatted.Length;
        for (int i = 0; i < padding; i++)
        {
            result.Append('0');
        }

        result.Append(formatted);
    }

    private static string FormatNumber(double value, DecimalSection format, CultureInfo culture)
    {
        bool thousandSeparator = format.ThousandSeparator;
        value = value / format.ThousandDivisor;
        value = value * format.PercentMultiplier;

        StringBuilder result = new();
        FormatNumber(
            value,
            format.BeforeDecimal,
            format.DecimalSeparator,
            format.AfterDecimal,
            thousandSeparator,
            culture,
            result
        );
        return result.ToString();
    }

    // Large enough for any double formatted with "F<digits>": worst case is a value near
    // double.MaxValue (~309 integer digits) plus a generous decimal precision allowance.
    private const int MaxFormattedDoubleLength = 512;

    private static void FormatNumber(
        double value,
        ReadOnlyMemory<char>[]? beforeDecimal,
        bool decimalSeparator,
        ReadOnlyMemory<char>[]? afterDecimal,
        bool thousandSeparator,
        CultureInfo culture,
        StringBuilder result
    )
    {
        int significantDigits = afterDecimal != null ? GetDigitCount(afterDecimal) : 0;

        Span<char> formatSpecBuffer = stackalloc char[12];
        formatSpecBuffer[0] = 'F';
        significantDigits.TryFormat(formatSpecBuffer[1..], out int specLength);
        ReadOnlySpan<char> formatSpec = formatSpecBuffer[..(1 + specLength)];

        Span<char> buffer = stackalloc char[MaxFormattedDoubleLength];
        if (
            !Math.Abs(value)
                .TryFormat(buffer, out int written, formatSpec, CultureInfo.InvariantCulture)
        )
        {
            // MaxFormattedDoubleLength comfortably covers every double (worst case ~309 integer
            // digits) at any decimal precision this library's format strings can express, so this
            // is unreachable in practice.
            throw new FormatException("Number format precision or magnitude is out of range.");
        }

        ReadOnlySpan<char> formatted = buffer[..written];
        int dotIndex = formatted.IndexOf('.');
        ReadOnlySpan<char> thousandsSpan = dotIndex >= 0 ? formatted[..dotIndex] : formatted;
        ReadOnlySpan<char> decimalSpan = dotIndex >= 0 ? formatted[(dotIndex + 1)..] : default;
        while (decimalSpan.Length > 0 && decimalSpan[^1] == '0')
        {
            decimalSpan = decimalSpan[..^1];
        }

        if (value < 0)
        {
            result.Append('-');
        }

        if (beforeDecimal != null)
        {
            FormatThousands(
                thousandsSpan,
                thousandSeparator,
                false,
                beforeDecimal,
                culture,
                result
            );
        }

        if (decimalSeparator)
        {
            result.Append(culture.NumberFormat.NumberDecimalSeparator);
        }

        if (afterDecimal != null)
        {
            FormatDecimals(decimalSpan, afterDecimal, result);
        }
    }

    /// <summary>
    /// Prints right-aligned, left-padded integer before the decimal separator. With optional most-significant zero.
    /// </summary>
    public static void FormatThousands(
        ReadOnlySpan<char> valueString,
        bool thousandSeparator,
        bool significantZero,
        ReadOnlyMemory<char>[] tokens,
        CultureInfo culture,
        StringBuilder result
    )
    {
        bool significant = false;
        int formatDigits = GetDigitCount(tokens);

        // Equivalent to valueString.PadLeft(formatDigits, '0') without allocating the padded
        // string: digitAt(i) reads a virtual left-zero-padded string of length paddedLength.
        int pad = Math.Max(0, formatDigits - valueString.Length);
        int paddedLength = valueString.Length + pad;

        // Print literals occurring before any placeholders
        int tokenIndex = 0;
        for (; tokenIndex < tokens.Length; tokenIndex++)
        {
            ReadOnlySpan<char> token = tokens[tokenIndex].Span;
            if (Token.IsPlaceholder(token))
            {
                break;
            }
            else
            {
                FormatLiteral(token, result);
            }
        }

        // Print value digits until there are as many digits remaining as there are placeholders
        int digitIndex = 0;
        for (; digitIndex < (paddedLength - formatDigits); digitIndex++)
        {
            significant = true;
            result.Append(digitIndex < pad ? '0' : valueString[digitIndex - pad]);

            if (thousandSeparator)
            {
                FormatThousandSeparator(paddedLength, digitIndex, culture, result);
            }
        }

        // Print remaining value digits and format literals
        for (; tokenIndex < tokens.Length; ++tokenIndex)
        {
            ReadOnlySpan<char> token = tokens[tokenIndex].Span;
            if (Token.IsPlaceholder(token))
            {
                char c = digitIndex < pad ? '0' : valueString[digitIndex - pad];
                if (c != '0' || (significantZero && digitIndex == paddedLength - 1))
                {
                    significant = true;
                }

                FormatPlaceholder(token, c, significant, result);

                if (thousandSeparator && (significant || token is "0"))
                {
                    FormatThousandSeparator(paddedLength, digitIndex, culture, result);
                }

                digitIndex++;
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
    }

    private static void FormatThousandSeparator(
        int valueLength,
        int digit,
        CultureInfo culture,
        StringBuilder result
    )
    {
        int positionInTens = valueLength - 1 - digit;
        if (positionInTens > 0 && (positionInTens % 3) == 0)
        {
            result.Append(culture.NumberFormat.NumberGroupSeparator);
        }
    }

    /// <summary>
    /// Prints left-aligned, right-padded integer after the decimal separator. Does not print significant zero.
    /// </summary>
    public static void FormatDecimals(
        ReadOnlySpan<char> valueString,
        ReadOnlyMemory<char>[] tokens,
        StringBuilder result
    )
    {
        bool significant = true;
        int unpaddedDigits = valueString.Length;

        // Equivalent to valueString.PadRight(formatDigits, '0') without allocating: positions past
        // the end of valueString simply read as '0'. GetDigitCount(tokens) == the number of
        // placeholder tokens, so valueIndex below never runs past that anyway.

        // Print all format digits
        int valueIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Length; ++tokenIndex)
        {
            ReadOnlySpan<char> token = tokens[tokenIndex].Span;
            if (Token.IsPlaceholder(token))
            {
                char c = valueIndex < valueString.Length ? valueString[valueIndex] : '0';
                significant = valueIndex < unpaddedDigits;

                FormatPlaceholder(token, c, significant, result);
                valueIndex++;
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
    }

    private static string FormatExponential(double value, Section format, CultureInfo culture)
    {
        // The application shall display a number to the right of
        // the "E" symbol that corresponds to the number of places that
        // the decimal point was moved.

        int baseDigits = 0;
        if (format.Exponential!.BeforeDecimal != null)
        {
            baseDigits = GetDigitCount(format.Exponential.BeforeDecimal);
        }

        int exponent = (int)Math.Floor(Math.Log10(Math.Abs(value)));
        double mantissa = value / Math.Pow(10, exponent);

        int shift = Math.Abs(exponent) % baseDigits;
        if (shift > 0)
        {
            if (exponent < 0)
            {
                shift = (baseDigits - shift);
            }

            mantissa *= Math.Pow(10, shift);
            exponent -= shift;
        }

        StringBuilder result = new();
        FormatNumber(
            mantissa,
            format.Exponential.BeforeDecimal,
            format.Exponential.DecimalSeparator,
            format.Exponential.AfterDecimal,
            false,
            culture,
            result
        );

        ReadOnlySpan<char> exponentialToken = format.Exponential.ExponentialToken.Span;
        result.Append(exponentialToken[0]);

        if (exponentialToken[1] == '+' && exponent >= 0)
        {
            result.Append('+');
        }
        else if (exponent < 0)
        {
            result.Append('-');
        }

        Span<char> exponentBuffer = stackalloc char[16];
        Math.Abs(exponent)
            .TryFormat(
                exponentBuffer,
                out int exponentWritten,
                default,
                CultureInfo.InvariantCulture
            );
        FormatThousands(
            exponentBuffer[..exponentWritten],
            false,
            false,
            format.Exponential.Power!,
            culture,
            result
        );
        return result.ToString();
    }

    private static string FormatFraction(double value, Section format, CultureInfo culture)
    {
        int integral = 0;
        int numerator,
            denominator;

        bool sign = value < 0;

        if (format.Fraction!.IntegerPart != null)
        {
            integral = (int)Math.Truncate(value);
            value = Math.Abs(value - integral);
        }

        if (format.Fraction.DenominatorConstant != 0)
        {
            denominator = format.Fraction.DenominatorConstant;
            double rr = Math.Round(value * denominator);
            double b = Math.Floor(rr / denominator);
            numerator = (int)(rr - b * denominator);
        }
        else
        {
            int denominatorDigits = Math.Min(GetDigitCount(format.Fraction.Denominator!), 7);
            GetFraction(
                value,
                (int)Math.Pow(10, denominatorDigits) - 1,
                out numerator,
                out denominator
            );
        }

        // Don't hide fraction if at least one zero in the numerator format
        int numeratorZeros = GetZeroCount(format.Fraction.Numerator!);
        bool hideFraction = (
            format.Fraction.IntegerPart != null && numerator == 0 && numeratorZeros == 0
        );

        StringBuilder result = new();

        if (sign)
        {
            result.Append('-');
        }

        // Print integer part with significant zero if fraction part is hidden
        if (format.Fraction.IntegerPart != null)
        {
            Span<char> integralBuffer = stackalloc char[16];
            Math.Abs(integral)
                .TryFormat(
                    integralBuffer,
                    out int integralWritten,
                    default,
                    CultureInfo.InvariantCulture
                );
            FormatThousands(
                integralBuffer[..integralWritten],
                false,
                hideFraction,
                format.Fraction.IntegerPart,
                culture,
                result
            );
        }

        Span<char> numeratorBuffer = stackalloc char[16];
        Math.Abs(numerator)
            .TryFormat(
                numeratorBuffer,
                out int numeratorWritten,
                default,
                CultureInfo.InvariantCulture
            );
        Span<char> denominatorBuffer = stackalloc char[16];
        denominator.TryFormat(
            denominatorBuffer,
            out int denominatorWritten,
            default,
            CultureInfo.InvariantCulture
        );

        StringBuilder fraction = new();

        FormatThousands(
            numeratorBuffer[..numeratorWritten],
            false,
            true,
            format.Fraction.Numerator!,
            culture,
            fraction
        );

        fraction.Append('/');

        if (format.Fraction.DenominatorPrefix != null)
        {
            FormatThousands(
                ReadOnlySpan<char>.Empty,
                false,
                false,
                format.Fraction.DenominatorPrefix,
                culture,
                fraction
            );
        }

        if (format.Fraction.DenominatorConstant != 0)
        {
            fraction.Append(format.Fraction.DenominatorConstant);
        }
        else
        {
            FormatDenominator(
                denominatorBuffer[..denominatorWritten],
                format.Fraction.Denominator!,
                fraction
            );
        }

        if (format.Fraction.DenominatorSuffix != null)
        {
            FormatThousands(
                ReadOnlySpan<char>.Empty,
                false,
                false,
                format.Fraction.DenominatorSuffix,
                culture,
                fraction
            );
        }

        if (hideFraction)
        {
            result.Append(' ', fraction.Length);
        }
        else
        {
            result.Append(fraction);
        }

        if (format.Fraction.FractionSuffix != null)
        {
            FormatThousands(
                ReadOnlySpan<char>.Empty,
                false,
                false,
                format.Fraction.FractionSuffix,
                culture,
                result
            );
        }

        return result.ToString();
    }

    // Adapted from ssf.js 'frac()' helper
    private static void GetFraction(double x, int d, out int nom, out int den)
    {
        int sgn = x < 0 ? -1 : 1;
        double B = x * sgn;
        double P_2 = 0.0;
        double P_1 = 1.0;
        double P = 0.0;
        double Q_2 = 1.0;
        double Q_1 = 0.0;
        double Q = 0.0;
        double A = Math.Floor(B);
        while (Q_1 < d)
        {
            A = Math.Floor(B);
            P = A * P_1 + P_2;
            Q = A * Q_1 + Q_2;
            if ((B - A) < 0.00000005)
            {
                break;
            }

            B = 1 / (B - A);
            P_2 = P_1;
            P_1 = P;
            Q_2 = Q_1;
            Q_1 = Q;
        }
        if (Q > d)
        {
            if (Q_1 > d)
            {
                Q = Q_2;
                P = P_2;
            }
            else
            {
                Q = Q_1;
                P = P_1;
            }
        }
        nom = (int)(sgn * P);
        den = (int)Q;
    }

    /// <summary>
    /// Prints left-aligned, left-padded fraction integer denominator.
    /// Assumes tokens contain only placeholders, valueString has fewer or equal number of digits as tokens.
    /// </summary>
    public static void FormatDenominator(
        ReadOnlySpan<char> valueString,
        ReadOnlyMemory<char>[] tokens,
        StringBuilder result
    )
    {
        int formatDigits = GetDigitCount(tokens);

        // Equivalent to valueString.PadLeft(formatDigits, '0') without allocating: digitAt(i)
        // reads a virtual left-zero-padded string of length paddedLength.
        int pad = Math.Max(0, formatDigits - valueString.Length);
        int paddedLength = valueString.Length + pad;

        bool significant = false;
        int valueIndex = 0;
        for (int tokenIndex = 0; tokenIndex < tokens.Length; ++tokenIndex)
        {
            ReadOnlySpan<char> token = tokens[tokenIndex].Span;
            char c;
            if (valueIndex < paddedLength)
            {
                c = GetLeftAlignedValueDigit(
                    token,
                    valueString,
                    pad,
                    paddedLength,
                    valueIndex,
                    significant,
                    out valueIndex
                );

                if (c != '0')
                {
                    significant = true;
                }
            }
            else
            {
                c = '0';
                significant = false;
            }

            FormatPlaceholder(token, c, significant, result);
        }
    }

    /// <summary>
    /// Returns the first digit from the virtual left-zero-padded value (<paramref name="pad"/>
    /// leading zeros followed by <paramref name="valueString"/>). If the token is '?' returns the
    /// first significant digit, or '0' if there are no significant digits. The out valueIndex
    /// parameter contains the offset to the next digit.
    /// </summary>
    private static char GetLeftAlignedValueDigit(
        ReadOnlySpan<char> token,
        ReadOnlySpan<char> valueString,
        int pad,
        int paddedLength,
        int startIndex,
        bool significant,
        out int valueIndex
    )
    {
        char c;
        valueIndex = startIndex;
        if (valueIndex < paddedLength)
        {
            c = valueIndex < pad ? '0' : valueString[valueIndex - pad];
            valueIndex++;

            if (c != '0')
            {
                significant = true;
            }

            if (token is "?" && !significant)
            {
                // Eat insignificant zeros to left align denominator
                while (valueIndex < paddedLength)
                {
                    c = valueIndex < pad ? '0' : valueString[valueIndex - pad];
                    valueIndex++;

                    if (c != '0')
                    {
                        significant = true;
                        break;
                    }
                }
            }
        }
        else
        {
            c = '0';
            significant = false;
        }

        return c;
    }

    private static void FormatPlaceholder(
        ReadOnlySpan<char> token,
        char c,
        bool significant,
        StringBuilder result
    )
    {
        if (token is "0")
        {
            if (significant)
            {
                result.Append(c);
            }
            else
            {
                result.Append('0');
            }
        }
        else if (token is "#")
        {
            if (significant)
            {
                result.Append(c);
            }
        }
        else if (token is "?")
        {
            if (significant)
            {
                result.Append(c);
            }
            else
            {
                result.Append(' ');
            }
        }
    }

    private static int GetDigitCount(ReadOnlyMemory<char>[] tokens)
    {
        int counter = 0;
        foreach (ReadOnlyMemory<char> token in tokens)
        {
            if (Token.IsPlaceholder(token.Span))
            {
                counter++;
            }
        }
        return counter;
    }

    private static int GetZeroCount(ReadOnlyMemory<char>[] tokens)
    {
        int counter = 0;
        foreach (ReadOnlyMemory<char> token in tokens)
        {
            if (token.Span is "0")
            {
                counter++;
            }
        }
        return counter;
    }

    private static void FormatLiteral(ReadOnlySpan<char> token, StringBuilder result)
    {
        if (token is ",")
        {
            // skip commas
        }
        else if (token.Length == 2 && (token[0] == '*' || token[0] == '\\'))
        {
            // TODO: * = repeat to fill cell
            result.Append(token[1]);
        }
        else if (token.Length == 2 && token[0] == '_')
        {
            result.Append(' ');
        }
        else if (token.Length > 0 && token[0] == '"')
        {
            result.Append(token.Slice(1, token.Length - 2));
        }
        else
        {
            result.Append(token);
        }
    }
}

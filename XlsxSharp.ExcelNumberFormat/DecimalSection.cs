namespace XlsxSharp.ExcelNumberFormat;

internal class DecimalSection
{
    public bool ThousandSeparator { get; set; }

    public double ThousandDivisor { get; set; }

    public double PercentMultiplier { get; set; }

    public ReadOnlyMemory<char>[]? BeforeDecimal { get; set; }

    public bool DecimalSeparator { get; set; }

    public ReadOnlyMemory<char>[]? AfterDecimal { get; set; }

    public static bool TryParse(ReadOnlyMemory<char>[] tokens, out DecimalSection? format)
    {
        if (
            Parser.ParseNumberTokens(
                tokens,
                0,
                out ReadOnlyMemory<char>[]? beforeDecimal,
                out bool decimalSeparator,
                out ReadOnlyMemory<char>[]? afterDecimal
            ) == tokens.Length
        )
        {
            bool thousandSeparator;
            double divisor = GetTrailingCommasDivisor(tokens, out thousandSeparator);
            double multiplier = GetPercentMultiplier(tokens);

            format = new DecimalSection
            {
                BeforeDecimal = beforeDecimal,
                DecimalSeparator = decimalSeparator,
                AfterDecimal = afterDecimal,
                PercentMultiplier = multiplier,
                ThousandDivisor = divisor,
                ThousandSeparator = thousandSeparator,
            };

            return true;
        }

        format = null;
        return false;
    }

    private static double GetPercentMultiplier(ReadOnlyMemory<char>[] tokens)
    {
        // If there is a percentage literal in the part list, multiply the result by 100
        foreach (ReadOnlyMemory<char> token in tokens)
        {
            if (token.Span is "%")
            {
                return 100;
            }
        }

        return 1;
    }

    private static double GetTrailingCommasDivisor(
        ReadOnlyMemory<char>[] tokens,
        out bool thousandSeparator
    )
    {
        // This parses all comma literals in the part list:
        // Each comma after the last digit placeholder divides the result by 1000.
        // If there are any other commas, display the result with thousand separators.
        bool hasLastPlaceholder = false;
        double divisor = 1.0;

        for (int j = 0; j < tokens.Length; j++)
        {
            int tokenIndex = tokens.Length - 1 - j;
            ReadOnlySpan<char> token = tokens[tokenIndex].Span;

            if (!hasLastPlaceholder)
            {
                if (Token.IsPlaceholder(token))
                {
                    // Each trailing comma multiplies the divisor by 1000
                    for (int k = tokenIndex + 1; k < tokens.Length; k++)
                    {
                        token = tokens[k].Span;
                        if (token is ",")
                        {
                            divisor *= 1000.0;
                        }
                        else
                        {
                            break;
                        }
                    }

                    // Continue scanning backwards from the last digit placeholder,
                    // but now look for a thousand separator comma
                    hasLastPlaceholder = true;
                }
            }
            else
            {
                if (token is ",")
                {
                    thousandSeparator = true;
                    return divisor;
                }
            }
        }

        thousandSeparator = false;
        return divisor;
    }
}

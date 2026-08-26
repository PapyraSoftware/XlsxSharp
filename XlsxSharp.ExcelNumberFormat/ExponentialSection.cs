namespace XlsxSharp.ExcelNumberFormat;

internal class ExponentialSection
{
    public ReadOnlyMemory<char>[]? BeforeDecimal { get; set; }

    public bool DecimalSeparator { get; set; }

    public ReadOnlyMemory<char>[]? AfterDecimal { get; set; }

    public ReadOnlyMemory<char> ExponentialToken { get; set; }

    public ReadOnlyMemory<char>[]? Power { get; set; }

    public static bool TryParse(ReadOnlyMemory<char>[] tokens, out ExponentialSection? format)
    {
        format = null;

        ReadOnlyMemory<char> exponentialToken;

        int partCount = Parser.ParseNumberTokens(
            tokens,
            0,
            out ReadOnlyMemory<char>[]? beforeDecimal,
            out bool decimalSeparator,
            out ReadOnlyMemory<char>[]? afterDecimal
        );

        if (partCount == 0)
        {
            return false;
        }

        int position = partCount;
        if (position < tokens.Length && Token.IsExponent(tokens[position].Span))
        {
            exponentialToken = tokens[position];
            position++;
        }
        else
        {
            return false;
        }

        format = new ExponentialSection
        {
            BeforeDecimal = beforeDecimal,
            DecimalSeparator = decimalSeparator,
            AfterDecimal = afterDecimal,
            ExponentialToken = exponentialToken,
            Power = tokens[position..],
        };

        return true;
    }
}

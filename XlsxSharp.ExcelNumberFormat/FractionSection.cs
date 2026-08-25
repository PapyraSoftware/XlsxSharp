using System.Text;

namespace XlsxSharp.ExcelNumberFormat;

internal class FractionSection
{
    public ReadOnlyMemory<char>[]? IntegerPart { get; set; }

    public ReadOnlyMemory<char>[]? Numerator { get; set; }

    public ReadOnlyMemory<char>[]? DenominatorPrefix { get; set; }

    public ReadOnlyMemory<char>[]? Denominator { get; set; }

    public int DenominatorConstant { get; set; }

    public ReadOnlyMemory<char>[]? DenominatorSuffix { get; set; }

    public ReadOnlyMemory<char>[]? FractionSuffix { get; set; }

    public static bool TryParse(ReadOnlyMemory<char>[] tokens, out FractionSection? format)
    {
        ReadOnlyMemory<char>[]? numeratorParts = null;
        ReadOnlyMemory<char>[]? denominatorParts = null;

        for (int i = 0; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> part = tokens[i].Span;
            if (part is "/")
            {
                numeratorParts = tokens[..i];
                i++;
                denominatorParts = tokens[i..];
                break;
            }
        }

        if (numeratorParts == null)
        {
            format = null;
            return false;
        }

        GetNumerator(
            numeratorParts,
            out ReadOnlyMemory<char>[]? integerPart,
            out ReadOnlyMemory<char>[] numeratorPart
        );

        if (
            !TryGetDenominator(
                denominatorParts!,
                out ReadOnlyMemory<char>[]? denominatorPrefix,
                out ReadOnlyMemory<char>[]? denominatorPart,
                out int denominatorConstant,
                out ReadOnlyMemory<char>[]? denominatorSuffix,
                out ReadOnlyMemory<char>[]? fractionSuffix
            )
        )
        {
            format = null;
            return false;
        }

        format = new FractionSection
        {
            IntegerPart = integerPart,
            Numerator = numeratorPart,
            DenominatorPrefix = denominatorPrefix,
            Denominator = denominatorPart,
            DenominatorConstant = denominatorConstant,
            DenominatorSuffix = denominatorSuffix,
            FractionSuffix = fractionSuffix,
        };

        return true;
    }

    private static void GetNumerator(
        ReadOnlyMemory<char>[] tokens,
        out ReadOnlyMemory<char>[]? integerPart,
        out ReadOnlyMemory<char>[] numeratorPart
    )
    {
        bool hasPlaceholder = false;
        bool hasSpace = false;
        bool hasIntegerPart = false;
        int numeratorIndex = -1;
        int index = tokens.Length - 1;
        while (index >= 0)
        {
            ReadOnlySpan<char> token = tokens[index].Span;
            if (Token.IsPlaceholder(token))
            {
                hasPlaceholder = true;

                if (hasSpace)
                {
                    hasIntegerPart = true;
                    break;
                }
            }
            else
            {
                if (hasPlaceholder && !hasSpace)
                {
                    // First time we get here marks the end of the integer part
                    hasSpace = true;
                    numeratorIndex = index + 1;
                }
            }
            index--;
        }

        if (hasIntegerPart)
        {
            integerPart = tokens[..numeratorIndex];
            numeratorPart = tokens[numeratorIndex..];
        }
        else
        {
            integerPart = null;
            numeratorPart = tokens;
        }
    }

    private static bool TryGetDenominator(
        ReadOnlyMemory<char>[] tokens,
        out ReadOnlyMemory<char>[]? denominatorPrefix,
        out ReadOnlyMemory<char>[]? denominatorPart,
        out int denominatorConstant,
        out ReadOnlyMemory<char>[]? denominatorSuffix,
        out ReadOnlyMemory<char>[]? fractionSuffix
    )
    {
        int index = 0;
        bool hasPlaceholder = false;
        bool hasConstant = false;

        StringBuilder constant = new();

        // Read literals until the first number placeholder or digit
        while (index < tokens.Length)
        {
            ReadOnlySpan<char> token = tokens[index].Span;
            if (Token.IsPlaceholder(token))
            {
                hasPlaceholder = true;
                break;
            }
            else if (Token.IsDigit19(token))
            {
                hasConstant = true;
                break;
            }
            index++;
        }

        if (!hasPlaceholder && !hasConstant)
        {
            denominatorPrefix = null;
            denominatorPart = null;
            denominatorConstant = 0;
            denominatorSuffix = null;
            fractionSuffix = null;
            return false;
        }

        // The denominator starts here, keep the index
        int denominatorIndex = index;

        // Read placeholders or digits in sequence
        while (index < tokens.Length)
        {
            ReadOnlySpan<char> token = tokens[index].Span;
            if (hasPlaceholder && Token.IsPlaceholder(token))
            {
                ; // OK
            }
            else if (hasConstant && (Token.IsDigit09(token)))
            {
                constant.Append(token);
            }
            else
            {
                break;
            }
            index++;
        }

        // 'index' is now at the first token after the denominator placeholders.
        // The remaining, if anything, is to be treated in one or two parts:
        // Any ultimately terminating literals are considered the "Fraction suffix".
        // Anything between the denominator and the fraction suffix is the "Denominator suffix".
        // Placeholders in the denominator suffix are treated as insignificant zeros.

        // Scan backwards to determine the fraction suffix
        int fractionSuffixIndex = tokens.Length;
        while (fractionSuffixIndex > index)
        {
            ReadOnlySpan<char> token = tokens[fractionSuffixIndex - 1].Span;
            if (Token.IsPlaceholder(token))
            {
                break;
            }

            fractionSuffixIndex--;
        }

        // Finally extract the detected token ranges

        denominatorPrefix = denominatorIndex > 0 ? tokens[..denominatorIndex] : null;

        if (hasConstant)
        {
            denominatorConstant = int.Parse(constant.ToString());
        }
        else
        {
            denominatorConstant = 0;
        }

        denominatorPart = tokens[denominatorIndex..index];

        denominatorSuffix = index < fractionSuffixIndex ? tokens[index..fractionSuffixIndex] : null;

        fractionSuffix = fractionSuffixIndex < tokens.Length ? tokens[fractionSuffixIndex..] : null;

        return true;
    }
}

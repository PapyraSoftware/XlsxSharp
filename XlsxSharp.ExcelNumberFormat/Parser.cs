using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace XlsxSharp.ExcelNumberFormat;

internal static class Parser
{
    public static List<Section> ParseSections(string formatString, out bool syntaxError)
    {
        Tokenizer tokenizer = new(formatString);
        List<Section> sections = new();
        syntaxError = false;
        while (true)
        {
            Section? section = ParseSection(tokenizer, sections.Count, out bool sectionSyntaxError);

            if (sectionSyntaxError)
            {
                syntaxError = true;
            }

            if (section == null)
            {
                break;
            }

            sections.Add(section);
        }

        return sections;
    }

    private static Section? ParseSection(Tokenizer reader, int index, out bool syntaxError)
    {
        bool hasDateParts = false;
        bool hasDurationParts = false;
        bool hasGeneralPart = false;
        bool hasTextPart = false;
        bool hasPlaceholders = false;
        Condition? condition = null;
        List<ReadOnlyMemory<char>> tokenList = new();

        syntaxError = false;
        while (TryReadToken(reader, out syntaxError, out ReadOnlyMemory<char> tokenMemory))
        {
            ReadOnlySpan<char> token = tokenMemory.Span;

            if (token is ";")
            {
                break;
            }

            hasPlaceholders |= Token.IsPlaceholder(token);

            if (Token.IsDatePart(token))
            {
                hasDateParts = true;
                hasDurationParts |= Token.IsDurationPart(token);
                tokenList.Add(tokenMemory);
            }
            else if (Token.IsGeneral(token))
            {
                hasGeneralPart = true;
                tokenList.Add(tokenMemory);
            }
            else if (token is "@")
            {
                hasTextPart = true;
                tokenList.Add(tokenMemory);
            }
            else if (token.Length > 0 && token[0] == '[')
            {
                // Does not add to tokens. Absolute/elapsed time tokens
                // also start with '[', but handled as date part above
                ReadOnlyMemory<char> expression = tokenMemory.Slice(1, tokenMemory.Length - 2);
                if (TryParseCondition(expression, out Condition? parseCondition))
                {
                    condition = parseCondition;
                }
                else if (TryParseCurrencySymbol(expression, out string? parseCurrencySymbol))
                {
                    tokenList.Add(("\"" + parseCurrencySymbol + "\"").AsMemory());
                }
            }
            else
            {
                tokenList.Add(tokenMemory);
            }
        }

        if (syntaxError || tokenList.Count == 0)
        {
            return null;
        }

        if (
            (hasDateParts && (hasGeneralPart || hasTextPart))
            || (hasGeneralPart && (hasDateParts || hasTextPart))
            || (hasTextPart && (hasGeneralPart || hasDateParts))
        )
        {
            // Cannot mix date, general and/or text parts
            syntaxError = true;
            return null;
        }

        ReadOnlyMemory<char>[] tokens = tokenList.ToArray();

        SectionType type;
        FractionSection? fraction = null;
        ExponentialSection? exponential = null;
        DecimalSection? number = null;
        ReadOnlyMemory<char>[]? generalTextDateDuration = null;

        if (hasDateParts)
        {
            type = hasDurationParts ? SectionType.Duration : SectionType.Date;

            generalTextDateDuration = ParseMilliseconds(tokens);
        }
        else if (hasGeneralPart)
        {
            type = SectionType.General;
            generalTextDateDuration = tokens;
        }
        else if (hasTextPart || !hasPlaceholders)
        {
            type = SectionType.Text;
            generalTextDateDuration = tokens;
        }
        else if (FractionSection.TryParse(tokens, out fraction))
        {
            type = SectionType.Fraction;
        }
        else if (ExponentialSection.TryParse(tokens, out exponential))
        {
            type = SectionType.Exponential;
        }
        else if (DecimalSection.TryParse(tokens, out number))
        {
            type = SectionType.Number;
        }
        else
        {
            // Unable to parse format string
            syntaxError = true;
            return null;
        }

        return new Section
        {
            Type = type,
            SectionIndex = index,
            Condition = condition,
            Fraction = fraction,
            Exponential = exponential,
            Number = number,
            GeneralTextDateDurationParts = generalTextDateDuration,
        };
    }

    /// <summary>
    /// Parses as many placeholders and literals needed to format a number with optional decimals.
    /// Returns number of tokens parsed, or 0 if the tokens didn't form a number.
    /// </summary>
    internal static int ParseNumberTokens(
        ReadOnlyMemory<char>[] tokens,
        int startPosition,
        out ReadOnlyMemory<char>[]? beforeDecimal,
        out bool decimalSeparator,
        out ReadOnlyMemory<char>[]? afterDecimal
    )
    {
        beforeDecimal = null;
        afterDecimal = null;
        decimalSeparator = false;

        // Upper-bounded by tokens.Length, so a single exactly-sized buffer replaces what would
        // otherwise be a growable list surviving only for the duration of this scan.
        ReadOnlyMemory<char>[] remainderBuffer = new ReadOnlyMemory<char>[tokens.Length];
        int remainderCount = 0;
        int index;
        for (index = 0; index < tokens.Length; ++index)
        {
            ReadOnlyMemory<char> token = tokens[index];
            ReadOnlySpan<char> tokenSpan = token.Span;
            if (tokenSpan is "." && beforeDecimal == null)
            {
                decimalSeparator = true;
                beforeDecimal = tokens[..index]; // TODO: why not remainder? has only valid tokens...

                remainderCount = 0;
            }
            else if (Token.IsNumberLiteral(tokenSpan))
            {
                remainderBuffer[remainderCount++] = token;
            }
            else if (tokenSpan.Length > 0 && tokenSpan[0] == '[')
            {
                // ignore
            }
            else
            {
                break;
            }
        }

        if (remainderCount > 0)
        {
            ReadOnlyMemory<char>[] remainder = remainderBuffer[..remainderCount];
            if (beforeDecimal != null)
            {
                afterDecimal = remainder;
            }
            else
            {
                beforeDecimal = remainder;
            }
        }

        return index;
    }

    private static ReadOnlyMemory<char>[] ParseMilliseconds(ReadOnlyMemory<char>[] tokens)
    {
        // if tokens form .0 through .000.., combine to single subsecond token. The result can only
        // be shorter than or equal to tokens (consecutive "0"s after a "." collapse into one), so a
        // single upper-bounded buffer replaces a growable list.
        ReadOnlyMemory<char>[] result = new ReadOnlyMemory<char>[tokens.Length];
        int count = 0;
        for (int i = 0; i < tokens.Length; i++)
        {
            ReadOnlySpan<char> token = tokens[i].Span;
            if (token is ".")
            {
                int zeros = 0;
                while (i + 1 < tokens.Length && tokens[i + 1].Span is "0")
                {
                    i++;
                    zeros++;
                }

                result[count++] =
                    zeros > 0 ? ("." + new string('0', zeros)).AsMemory() : ".".AsMemory();
            }
            else
            {
                result[count++] = tokens[i];
            }
        }

        return count == result.Length ? result : result[..count];
    }

    private static bool TryReadToken(
        Tokenizer reader,
        out bool syntaxError,
        out ReadOnlyMemory<char> token
    )
    {
        int offset = reader.Position;
        if (
            ReadLiteral(reader)
            || reader.ReadEnclosed('[', ']')
            ||
            // Symbols
            reader.ReadOneOf("#?,!&%+-$€£0123456789{}():;/.@ ")
            || reader.ReadString("e+", true)
            || reader.ReadString("e-", true)
            || reader.ReadString("General", true)
            ||
            // Date
            reader.ReadString("am/pm", true)
            || reader.ReadString("a/p", true)
            || reader.ReadOneOrMore('y')
            || reader.ReadOneOrMore('Y')
            || reader.ReadOneOrMore('m')
            || reader.ReadOneOrMore('M')
            || reader.ReadOneOrMore('d')
            || reader.ReadOneOrMore('D')
            || reader.ReadOneOrMore('h')
            || reader.ReadOneOrMore('H')
            || reader.ReadOneOrMore('s')
            || reader.ReadOneOrMore('S')
            || reader.ReadOneOrMore('g')
            || reader.ReadOneOrMore('G')
        )
        {
            syntaxError = false;
            int length = reader.Position - offset;
            token = reader.Slice(offset, length);
            return true;
        }

        syntaxError = reader.Position < reader.Length;
        token = default;
        return false;
    }

    private static bool ReadLiteral(Tokenizer reader)
    {
        if (reader.Peek() == '\\' || reader.Peek() == '*' || reader.Peek() == '_')
        {
            reader.Advance(2);
            return true;
        }
        else if (reader.ReadEnclosed('"', '"'))
        {
            return true;
        }

        return false;
    }

    private static bool TryParseCondition(
        ReadOnlyMemory<char> token,
        [NotNullWhen(true)] out Condition? result
    )
    {
        Tokenizer tokenizer = new(token);

        if (
            tokenizer.ReadString("<=")
            || tokenizer.ReadString("<>")
            || tokenizer.ReadString("<")
            || tokenizer.ReadString(">=")
            || tokenizer.ReadString(">")
            || tokenizer.ReadString("=")
        )
        {
            int conditionPosition = tokenizer.Position;
            string op = tokenizer.Slice(0, conditionPosition).ToString();

            if (ReadConditionValue(tokenizer))
            {
                ReadOnlySpan<char> valueSpan = tokenizer
                    .Slice(conditionPosition, tokenizer.Position - conditionPosition)
                    .Span;

                result = new Condition
                {
                    Operator = op,
                    Value = double.Parse(valueSpan, CultureInfo.InvariantCulture),
                };
                return true;
            }
        }

        result = null;
        return false;
    }

    private static bool ReadConditionValue(Tokenizer tokenizer)
    {
        // NFPartCondNum = [ASCII-HYPHEN-MINUS] NFPartIntNum [INTL-CHAR-DECIMAL-SEP NFPartIntNum] [NFPartExponential NFPartIntNum]
        tokenizer.ReadString("-");
        while (tokenizer.ReadOneOf("0123456789")) { }

        if (tokenizer.ReadString("."))
        {
            while (tokenizer.ReadOneOf("0123456789")) { }
        }

        if (tokenizer.ReadString("e+", true) || tokenizer.ReadString("e-", true))
        {
            if (tokenizer.ReadOneOf("0123456789"))
            {
                while (tokenizer.ReadOneOf("0123456789")) { }
            }
            else
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryParseCurrencySymbol(
        ReadOnlyMemory<char> token,
        [NotNullWhen(true)] out string? currencySymbol
    )
    {
        ReadOnlySpan<char> span = token.Span;
        if (span.IsEmpty || span[0] != '$')
        {
            currencySymbol = null;
            return false;
        }

        int dashIndex = span.IndexOf('-');
        currencySymbol =
            dashIndex >= 0 ? span.Slice(1, dashIndex - 1).ToString() : span[1..].ToString();

        return true;
    }
}

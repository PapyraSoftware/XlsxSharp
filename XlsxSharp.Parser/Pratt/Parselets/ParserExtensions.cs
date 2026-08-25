using static XlsxSharp.Parser.Pratt.CompatUtils;

namespace XlsxSharp.Parser.Pratt.Parselets;

internal static class ParserExtensions
{
    private const int MIN_A1_LENGTH = 2; // A1
    private const int MAX_A1_LENGTH = 1 + 3 + 1 + 7; // $XFD$1048576
    private const int MIN_COL_LENGTH = 1; // A
    private const int MAX_COL_LENGTH = 4; // $XFD
    private const int MIN_ROW_LENGTH = 1; // 1
    private const int MAX_ROW_LENGTH = 8; // $1048576

    /// <summary>
    /// Parse a function call argument list, having already consumed the opening
    /// <see cref="TokenType.LeftParen"/>. Arguments may be blank (e.g. <c>SUM(1,,2)</c>,
    /// <c>SUM(1,)</c>, <c>SUM(,1)</c>), but an entirely empty list (<c>SUM()</c>) has zero
    /// arguments rather than a single blank one.
    /// </summary>
    public static (List<T> Args, Token RightParen) ParseArgumentList<TScalar, T, TContext>(this Parser<T, TContext> parser, IAstFactory<TScalar, T, TContext> factory, TContext ctx)
    {
        List<T> args = [];

        // Unlike every other whitespace check in the argument list (all reached only after
        // ParseExpression's own loop already consumed any preceding whitespace as a side effect
        // of looking for an operator), this one runs before any argument has been parsed, so nothing
        // has skipped a leading "( )" gap yet.
        parser.SkipWhitespace();
        if (parser.LookAhead(1).Type == TokenType.RightParen)
        {
            return (args, parser.Consume(TokenType.RightParen));
        }

        while (true)
        {
            Token next = parser.LookAhead(1);
            if (next.Type == TokenType.Comma)
            {
                Token comma = parser.Consume(TokenType.Comma);
                args.Add(factory.BlankNode(ctx, new SymbolRange(comma.Range.Start, comma.Range.Start)));
                continue;
            }

            if (next.Type == TokenType.RightParen)
            {
                Token rightParen = parser.Consume(TokenType.RightParen);
                args.Add(factory.BlankNode(ctx, new SymbolRange(rightParen.Range.Start, rightParen.Range.Start)));
                return (args, rightParen);
            }

            // A top-level "," inside an argument means "next argument" here, not "union" - see
            // Parser.SkipUnion. A nested "(...)" still allows union for its own content (e.g.
            // SUM((A1,B2))), since GroupParselet resets SkipUnion for whatever it parses.
            bool previousSkipUnion = parser.SkipUnion;
            parser.SkipUnion = true;
            Node<T> arg = parser.ParseExpression(ctx, 0);
            parser.SkipUnion = previousSkipUnion;
            args.Add(arg.Value);

            if (parser.LookAhead(1).Type == TokenType.RightParen)
            {
                return (args, parser.Consume(TokenType.RightParen));
            }

            parser.Consume(TokenType.Comma);
        }
    }

    /// <summary>
    /// Strip the surrounding double quotes of a <see cref="TokenType.Text"/> token and collapse
    /// escaped <c>""</c> pairs into a single <c>"</c>, e.g. <c>"a""b"</c> becomes <c>a"b</c>.
    /// </summary>
    public static string UnescapeText(ReadOnlySpan<char> quotedText)
    {
        ReadOnlySpan<char> inner = quotedText[1..^1];
        if (inner.IndexOf('"') < 0)
        {
            return inner.ToString();
        }

        Span<char> buffer = new char[inner.Length];
        int w = 0;
        int i = 0;
        while (i < inner.Length)
        {
            if (inner[i] == '"')
            {
                i++;
            }

            buffer[w++] = inner[i++];
        }

        return buffer[..w].ToString();
    }

    public static bool EqualCaseInsensitive(ReadOnlySpan<char> text, string other)
    {
        if (text.Length != other.Length)
        {
            return false;
        }

        return text.CompareTo(other.AsSpan(), StringComparison.OrdinalIgnoreCase) == 0;
    }

    /// <summary>
    /// The functions the oracle's lexer recognizes (via a dedicated REF_FUNCTION_LIST token,
    /// distinct from every other function name) as capable of returning a reference, and therefore
    /// as a valid operand of the range operator (<c>:</c>) - but only called unqualified: even
    /// <c>Sheet1!INDEX(...)</c> doesn't count, since sheet-qualified calls are always parsed
    /// through the plain (never reference-shaped) function-call grammar production instead.
    /// </summary>
    private static readonly string[] RefFunctionNames = ["CHOOSE", "IF", "INDEX", "INDIRECT", "OFFSET"];

    public static bool IsRefFunctionName(ReadOnlySpan<char> name)
    {
        foreach (string candidate in RefFunctionNames)
        {
            if (EqualCaseInsensitive(name, candidate))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Peek past a <see cref="TokenType.Range"/> (<c>:</c>) starting at lookahead distance 1,
    /// tolerating up to one insignificant <see cref="TokenType.Whitespace"/> token immediately
    /// before and/or after it - matching the oracle's COLON lexer token, whose own regex absorbs
    /// such whitespace directly into the token. Pure lookahead: nothing is consumed, so a caller
    /// can use this to decide whether to commit to consuming the pattern at all.
    /// </summary>
    private static bool TryPeekRangeContinuation<T, TContext>(this Parser<T, TContext> parser, out Token afterColon, out int distance)
    {
        distance = 1;
        Token maybeColon = parser.LookAhead(distance);
        if (maybeColon.Type == TokenType.Whitespace)
        {
            maybeColon = parser.LookAhead(++distance);
        }

        if (maybeColon.Type != TokenType.Range)
        {
            afterColon = default;
            return false;
        }

        Token next = parser.LookAhead(++distance);
        if (next.Type == TokenType.Whitespace)
        {
            next = parser.LookAhead(++distance);
        }

        afterColon = next;
        return true;
    }

    /// <summary>
    /// Having confirmed (via <see cref="TryPeekRangeContinuation{T,TContext}"/>) that a range
    /// continuation is present, actually consume it: an optional whitespace, the range operator,
    /// another optional whitespace, then the second-corner token itself.
    /// </summary>
    private static Token ConsumeRangeContinuation<T, TContext>(this Parser<T, TContext> parser)
    {
        parser.SkipWhitespace();
        parser.Consume(TokenType.Range);
        parser.SkipWhitespace();
        return parser.Consume();
    }

    public static bool TryReferenceA1<T, TContext>(this Parser<T, TContext> parser, Token token, out ReferenceArea area, out SymbolRange range)
    {
        if (token.Type is not TokenType.Ident and not TokenType.Number)
        {
            area = default;
            range = default;
            return false;
        }

        // Check for area `A1:B2` or just cell `A1`
        if (parser.TryLocalAreaA1(token, out area, out range))
        {
            return true;
        }

        // Check for colspan `A:B`
        if (parser.TryLocalColSpanA1(token, out area, out range))
        {
            return true;
        }

        // Check for rowspan `1:2`, can be ident or number token
        if (parser.TryLocalRowSpanA1(token, out area, out range))
        {
            return true;
        }

        return false;
    }

    public static bool TryLocalAreaA1<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReferenceArea area, out SymbolRange range)
    {
        if (identToken.Type != TokenType.Ident)
        {
            area = default;
            range = default;
            return false;
        }

        ReadOnlySpan<char> ident = identToken.GetText(parser.Input);

        if (TryGetCellA1(ident, out RowCol cell1))
        {
            if (parser.TryPeekRangeContinuation(out Token maybeCell2Token, out _) &&
                maybeCell2Token.Type == TokenType.Ident &&
                TryGetCellA1(maybeCell2Token.GetText(parser.Input), out RowCol cell2))
            {
                // Result: area A1:B2
                // The code is joining two cells into an area through range operator, but that
                // is allowed. Range is highest priority operator, left to right associativity.
                Token cell2Token = parser.ConsumeRangeContinuation();

                area = new ReferenceArea(cell1, cell2);
                range = new SymbolRange(identToken.Range.Start, cell2Token.Range.End);
                return true;
            }

            // Result: cell A1
            area = new ReferenceArea(cell1);
            range = identToken.Range;
            return true;
        }

        range = default;
        area = default;
        return false;
    }

    public static bool TryLocalColSpanA1<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReferenceArea area, out SymbolRange range)
    {
        if (identToken.Type != TokenType.Ident)
        {
            area = default;
            range = default;
            return false;
        }

        ReadOnlySpan<char> ident = identToken.GetText(parser.Input);

        // Careful, 'A' can be just a name without the other column
        if (TryGetColA1(ident, out RowCol col1) &&
            parser.TryPeekRangeContinuation(out Token maybeCol2Token, out int distance) &&
            maybeCol2Token.Type == TokenType.Ident &&
            TryGetColA1(maybeCol2Token.GetText(parser.Input), out RowCol col2) &&
            // "Jan:Dec!A1" is a 3D sheet range reference (IdentParselet has its own dedicated
            // check for that), not a column span "JAN:DEC" that happens to be followed by
            // something else - column letters and sheet names are lexically indistinguishable, so
            // this is a genuine ambiguity the oracle's lexer resolves by never treating text
            // immediately followed by "!" as a bare A1_SPAN_REFERENCE token in the first place.
            parser.LookAhead(distance + 1).Type != TokenType.Bang)
        {
            // Result: colspan A:B
            Token col2Token = parser.ConsumeRangeContinuation();

            area = new ReferenceArea(col1, col2);
            range = new SymbolRange(identToken.Range.Start, col2Token.Range.End);
            return true;
        }

        area = default;
        range = default;
        return false;
    }

    public static bool TryLocalRowSpanA1<T, TContext>(this Parser<T, TContext> parser, Token numberOrIdentToken, out ReferenceArea area, out SymbolRange range)
    {
        if (numberOrIdentToken.Type is not TokenType.Ident and not TokenType.Number)
        {
            area = default;
            range = default;
            return false;
        }

        ReadOnlySpan<char> numberOrIdent = numberOrIdentToken.GetText(parser.Input);

        if (TryGetRowA1(numberOrIdent, out RowCol row1) &&
            parser.TryPeekRangeContinuation(out Token maybeRow2Token, out int distance) &&
            maybeRow2Token.Type is TokenType.Number or TokenType.Ident &&
            TryGetRowA1(maybeRow2Token.GetText(parser.Input), out RowCol row2) &&
            // Same "not a 3D sheet range" guard as TryLocalColSpanA1 - see the comment there.
            parser.LookAhead(distance + 1).Type != TokenType.Bang)
        {
            // Result: rowspan 1:2
            Token row2Token = parser.ConsumeRangeContinuation();

            area = new ReferenceArea(row1, row2);
            range = new SymbolRange(numberOrIdentToken.Range.Start, row2Token.Range.End);
            return true;
        }

        area = default;
        range = default;
        return false;
    }

    public static bool TryGetUnquotedSheet<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReadOnlySpan<char> sheetName)
    {
        ReadOnlySpan<char> text = identToken.GetText(parser.Input);
        bool isUnquotedSheet = NameUtils.IsSheetNameValid(text) && !NameUtils.ShouldQuote(text);
        if (isUnquotedSheet)
        {
            sheetName = text;
            return true;
        }

        sheetName = default;
        return false;
    }

    public static bool TryGetName<T, TContext>(this Parser<T, TContext> parser, Token identToken, out ReadOnlySpan<char> name)
    {
        if (identToken.Type != TokenType.Ident)
        {
            name = default;
            return false;
        }

        ReadOnlySpan<char> text = identToken.GetText(parser.Input);

        // TRUE/FALSE are never a valid NAME token, even though the text alone matches the name
        // grammar - the oracle's lexer always classifies them as LOGICAL_CONSTANT instead (unless
        // immediately followed by "(", where they lex as a function name; that path never calls
        // TryGetName, so it isn't affected here).
        if (
            NameUtils.IsNameValid(text)
            && !EqualCaseInsensitive(text, "TRUE")
            && !EqualCaseInsensitive(text, "FALSE")
        )
        {
            name = text;
            return true;
        }

        name = default;
        return false;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid A1 cell reference? No padding, case insensitive.
    /// </summary>
    public static bool TryGetCellA1(ReadOnlySpan<char> text, out RowCol cell)
    {
        cell = default;
        if (text.Length is < MIN_A1_LENGTH or > MAX_A1_LENGTH)
        {
            return false;
        }

        int i = 0;
        bool absCol = text[i] == '$';
        if (absCol)
        {
            ++i;
        }

        int col = 0;
        while (i < text.Length && IsAsciiLetter(text[i]))
        {
            col = col * 26 + GetColIndex(text[i++]) + 1;
        }

        if (col is < RowCol.MinCol or > RowCol.MaxCol || i >= text.Length)
        {
            return false;
        }

        bool absRow = text[i] == '$';
        if (absRow)
        {
            if (++i >= text.Length)
            {
                return false;
            }
        }

        if (text[i] == '0')
        {
            return false;
        }

        int row = 0;
        while (i < text.Length && IsAsciiDigit(text[i]))
        {
            row = row * 10 + text[i++] - '0';
        }

        if (row is < RowCol.MinRow or > RowCol.MaxRow || i < text.Length)
        {
            return false;
        }

        cell = new RowCol(
            absRow ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, row,
            absCol ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, col,
            A1);
        return true;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid end of an A1 colspan? No padding, case insensitive.
    /// Valid examples: <c>A</c>, <c>a</c>, <c>$A</c>, <c>$XFD</c>.
    /// Invalid examples: <c> A </c>, <c>$ a</c>, <c>$</c>, <c>$XFE</c>.
    /// </summary>
    public static bool TryGetColA1(ReadOnlySpan<char> text, out RowCol colRef)
    {
        colRef = default;
        if (text.Length is < MIN_COL_LENGTH or > MAX_COL_LENGTH)
        {
            return false;
        }

        int i = 0;
        bool absCol = text[i] == '$';
        if (absCol)
        {
            ++i;
        }

        int col = 0;
        while (i < text.Length && IsAsciiLetter(text[i]))
        {
            col = col * 26 + GetColIndex(text[i++]) + 1;
        }

        if (col is < RowCol.MinCol or > RowCol.MaxCol || i < text.Length)
        {
            return false;
        }

        colRef = new RowCol(
            ReferenceAxisType.None, 0,
            absCol ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, col,
            A1);
        return true;
    }

    /// <summary>
    /// Is the <paramref name="text"/> a valid end of an A1 rowspan? No padding.
    /// Valid examples: <c>1</c>, <c>$1</c>, <c>$1048576</c>.
    /// Invalid examples: <c>1.0</c>, <c>$ 1</c>, <c>$</c>, <c>$1048577</c>.
    /// </summary>
    public static bool TryGetRowA1(ReadOnlySpan<char> text, out RowCol rowRef)
    {
        rowRef = default;
        if (text.Length is < MIN_ROW_LENGTH or > MAX_ROW_LENGTH)
        {
            return false;
        }

        int i = 0;
        bool absRow = text[i] == '$';
        if (absRow)
        {
            if (++i >= text.Length)
            {
                return false;
            }
        }

        if (text[i] == '0')
        {
            return false;
        }

        int row = 0;
        while (i < text.Length && IsAsciiDigit(text[i]))
        {
            row = row * 10 + text[i++] - '0';
        }

        if (row is < RowCol.MinRow or > RowCol.MaxRow || i < text.Length)
        {
            return false;
        }

        rowRef = new RowCol(
            absRow ? ReferenceAxisType.Absolute : ReferenceAxisType.Relative, row,
            ReferenceAxisType.None, 0,
            A1);
        return true;
    }

    private static int GetColIndex(char asciiLetter)
    {
        return (asciiLetter | 0x20) - 'a';
    }
}

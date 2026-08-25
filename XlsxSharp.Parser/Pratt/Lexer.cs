using System.Xml;

namespace XlsxSharp.Parser.Pratt;

/// <summary>
/// A lexer for pratt parser.
/// </summary>
internal class Lexer
{
    private const int EOF = -1;

    // Every single-char operator that doesn't need lookahead to disambiguate (unlike '<'/'>',
    // which may start '<=', '<>' or '>=') maps directly to its TokenType here, so Next() can
    // dispatch it with one array lookup instead of a chain of up to 17 sequential comparisons.
    private static readonly TokenType?[] SingleCharOperator;

    // A small ring buffer of already-lexed lookahead tokens, indexed directly by Peek instead of
    // walked with an enumerator like a BCL Queue<T> would require - PeekPastWhitespace (see Parser)
    // calls Peek up to twice per reference atom, the most common leaf in real formulas, so this is
    // on the hottest path in the parser.
    private Token[] _lookahead = new Token[4];
    private int _lookaheadHead;
    private int _lookaheadCount;
    private string _input = string.Empty; // Currently tokenized formula
    private int _start; // The start index of currently parsed token in Next()
    private int _i; // Index of current code point _c in _input
    private int _c; // A current code point (including astral planes) or -1 if at the EOF
    private bool _isR1C1; // True while tokenizing an R1C1-style formula rather than A1


    static Lexer()
    {
        SingleCharOperator = new TokenType?[128];
        SingleCharOperator['+'] = TokenType.Plus;
        SingleCharOperator['-'] = TokenType.Minus;
        SingleCharOperator['*'] = TokenType.Mul;
        SingleCharOperator['/'] = TokenType.Div;
        SingleCharOperator['^'] = TokenType.Pow;
        SingleCharOperator['%'] = TokenType.Percent;
        SingleCharOperator['&'] = TokenType.Concat;
        SingleCharOperator['!'] = TokenType.Bang;
        SingleCharOperator['('] = TokenType.LeftParen;
        SingleCharOperator[')'] = TokenType.RightParen;
        SingleCharOperator['{'] = TokenType.LeftCurly;
        SingleCharOperator['}'] = TokenType.RightCurly;
        SingleCharOperator[','] = TokenType.Comma;
        SingleCharOperator[';'] = TokenType.Semicolon;
        SingleCharOperator[':'] = TokenType.Range;
        SingleCharOperator['@'] = TokenType.Intersection;
        SingleCharOperator['='] = TokenType.Equal;
    }

    public Lexer()
        : this(string.Empty)
    {
    }

    public Lexer(string input, bool isR1C1 = false)
    {
        this.Reset(input, isR1C1);
    }

    private bool IsEof => this._c == EOF;

    /// <summary>
    /// Prepare lexer to start tokenization of the <paramref name="formula"/>.
    /// </summary>
    /// <param name="formula">Formula to tokenize.</param>
    /// <param name="isR1C1">Tokenize <paramref name="formula"/> as R1C1 rather than A1.</param>
    public void Reset(string formula, bool isR1C1 = false)
    {
        this._input = formula ?? throw new ArgumentNullException();
        this._start = -1;
        this._i = -1;
        this._c = 0;
        this._lookaheadHead = 0;
        this._lookaheadCount = 0;
        this._isR1C1 = isR1C1;
    }

    public Token Consume()
    {
        if (this._lookaheadCount == 0)
        {
            return this.Next();
        }

        Token token = this._lookahead[this._lookaheadHead];
        this._lookaheadHead = (this._lookaheadHead + 1) % this._lookahead.Length;
        this._lookaheadCount--;
        return token;
    }

    public Token Peek(int distance = 1)
    {
        while (this._lookaheadCount < distance)
        {
            if (this._lookaheadCount == this._lookahead.Length)
            {
                this.GrowLookahead();
            }

            int writeIndex = (this._lookaheadHead + this._lookaheadCount) % this._lookahead.Length;
            this._lookahead[writeIndex] = this.Next();
            this._lookaheadCount++;
        }

        int readIndex = (this._lookaheadHead + distance - 1) % this._lookahead.Length;
        return this._lookahead[readIndex];
    }

    private void GrowLookahead()
    {
        Token[] grown = new Token[this._lookahead.Length * 2];
        for (int i = 0; i < this._lookaheadCount; i++)
        {
            grown[i] = this._lookahead[(this._lookaheadHead + i) % this._lookahead.Length];
        }

        this._lookahead = grown;
        this._lookaheadHead = 0;
    }

    private Token Next()
    {
        if (this._i < 0)
        {
            this.Advance();
        }

        if (this.IsEof)
        {
            return new Token(TokenType.Eof, 0, 0);
        }

        this._start = this._i;

        // Number
        if (IsDigit(this._c))
        {
            // Whole number part
            DigitSequence();

            // Fractional part
            if (this._c == '.')
            {
                this.Advance();
                DigitSequence();
            }

            ExponentPart();

            return this.T(TokenType.Number);
        }
        if (this._c is '.')
        {
            this.Advance();

            // Fractional part
            DigitSequence();
            ExponentPart();

            return this.T(TokenType.Number);
        }

        // Text
        if (this._c == '"')
        {
            this.Advance();

            while (!this.IsEof)
            {
                if (this._c == '"')
                {
                    this.Advance();
                    if (this._c != '"')
                    {
                        return this.T(TokenType.Text);
                    }
                }

                if (!IsXml10Char(this._c))
                {
                    throw new ParsingException($"Invalid text character (codepoint {this._c:x8}).");
                }

                this.Advance();
            }

            throw ParsingException.UnterminatedLiteral(this._start, '"');
        }

        // QIdent
        if (this._c == '\'')
        {
            while (!this.IsEof)
            {
                this.Advance();

                if (this._c == '\'')
                {
                    this.Advance();
                    if (this._c != '\'')
                    {
                        return this.T(TokenType.QIdent);
                    }
                }
            }

            throw ParsingException.UnterminatedLiteral(this._start, '\'');
        }

        if (IsIdentStart(this._c))
        {
            if (this._isR1C1 && this._c is 'R' or 'r' or 'C' or 'c')
            {
                Token? r1c1Token = this.TryScanR1C1Reference();
                if (r1c1Token is { } found)
                {
                    return found;
                }

                // Not R1C1-shaped after all (e.g. "Revenue", "Costs", "R1C1style") - the speculative
                // scan above is guaranteed not to have consumed anything on failure, so fall through
                // to plain identifier scanning from the same, still-current position.
            }

            this.Advance();
            while (!this.IsEof && IsIdentNext(this._c))
            {
                this.Advance();
            }

            return this.T(TokenType.Ident);
        }

        if (this._c < SingleCharOperator.Length)
        {
            TokenType? op = SingleCharOperator[this._c];
            if (op is not null)
            {
                return FoundToken(op.Value);
            }
        }

        if (this._c == '<')
        {
            int next = this.Advance();
            if (next == '>')
            {
                return FoundToken(TokenType.NotEqual);
            }

            if (next == '=')
            {
                return FoundToken(TokenType.LessEqual);
            }

            return this.T(TokenType.Less);
        }

        if (this._c == '>')
        {
            if (this.Advance() == '=')
            {
                return FoundToken(TokenType.GreaterEqual);
            }

            return this.T(TokenType.Greater);
        }

        if (IsWhitespace(this._c))
        {
            do
            {
                this.Advance();
            } while (IsWhitespace(this._c));

            return this.T(TokenType.Whitespace);
        }

        // Spill operator and errors
        if (this._c == '#')
        {
            int char1 = this.Advance();
            switch (char1)
            {
                case 'D' or 'd':
                    return Error("#DIV/0!", 2);
                case 'R' or 'r':
                    return Error("#REF!", 2);
                case 'V' or 'v':
                    return Error("#VALUE!", 2);
                case 'G' or 'g':
                    return Error("#GETTING_DATA", 2);
                case 'N' or 'n':
                    {
                        int char2 = this.Advance();
                        if (char2 == '/')
                        {
                            return Error("#N/A", 3);
                        }

                        if (char2 is 'A' or 'a')
                        {
                            return Error("#NAME?", 3);
                        }

                        int char3 = this.Advance();
                        if ((char2 is 'U' or 'u') && (char3 is 'L' or 'l'))
                        {
                            return Error("#NULL!", 4);
                        }

                        if ((char2 is 'U' or 'u') && (char3 is 'M' or 'm'))
                        {
                            return Error("#NUM!", 4);
                        }

                        throw ParsingException.TokenPartialMatch(this._start, TokenType.Error);
                    }
            }

            return this.T(TokenType.Spill);
        }

        if (this._c == '[')
        {
            int level = 0;
            do
            {
                switch (this._c)
                {
                    case '[':
                        ++level;
                        break;
                    case ']':
                        --level;
                        break;
                    case '\'':
                        this.Advance(); // Escaped chars don't change level - skip
                        break;
                }

                if (this.IsEof)
                {
                    throw new ParsingException($"Unable to find closing square bracket for token from position {this._start}.");
                }

                if (level > 2)
                {
                    throw new ParsingException($"There can be at most two nested square brackets in a token from position {this._start}.");
                }

                this.Advance();
            } while (level > 0);

            return this.T(TokenType.SquareIdent);
        }

        throw ParsingException.UnableToSelectToken(this._start);

        static bool IsWhitespace(int c)
        {
            return c is ' ' or '\r' or '\n' or '\t';
        }

        static bool IsDigit(int c)
        {
            return c is >= '0' and <= '9';
        }

        // Check [0-9]+
        void DigitSequence()
        {
            do
            {
                if (!IsDigit(this._c))
                {
                    throw ParsingException.TokenPartialMatch(this._start, TokenType.Number);
                }

                this.Advance();
            }
            while (!this.IsEof && IsDigit(this._c));
        }

        void ExponentPart()
        {
            if (this._c is 'e' or 'E')
            {
                if (this.Advance() is '+' or '-')
                {
                    this.Advance();
                }

                DigitSequence();
            }
        }

        static bool IsIdentStart(int c)
        {
            // Ident must satisfy logical-literal, sheet-name, name and A1-cell/column/row.
            // The oracle's own NAME production (see LexerA1.rl) accepts *any* codepoint above
            // 0x7F here, not just letters - e.g. "‰" (U+2030, PER MILLE SIGN, a symbol) is just
            // as valid a NAME character as "é" is.
            return
                IsAsciiLetter(c) || // name + A1
                c == '$' || // A1
                (c is '_' or '\\' or '?') || // name
                c > 0x7F; // name
        }

        static bool IsIdentNext(int c)
        {
            return IsIdentStart(c) ||
                   c is >= '0' and <= '9' ||  // name, A1
                   c == '.'; // name + future-functions
        }

        Token Error(string error, int start)
        {
            foreach (char errorChar in error.AsSpan().Slice(start))
            {
                this.Advance();
                if (ToUpperAlpha(this._c) != errorChar)
                {
                    throw ParsingException.TokenPartialMatch(this._start, TokenType.Error);
                }
            }

            this.Advance();
            return this.T(TokenType.Error);
        }

        // Token that ends at the Current has been found. Advance to next and return token.
        Token FoundToken(TokenType type)
        {
            this.Advance();
            return this.T(type);
        }

        // Convert a-z to A-Z, keep other codepoints same.
        static int ToUpperAlpha(int codepoint)
        {
            return codepoint is >= 'a' and <= 'z'
                ? 'A' + codepoint - 'a'
                : codepoint;
        }

        static bool IsAsciiLetter(int codepoint)
        {
            // Convert to lowercase, normalize 'a' to 0, and check if within 0 (~A)..25(~Z).
            // Really cool use of cast int to uint (-1 to 0xFFFFFFFF), thus saving one comparison
            // and avoiding potential pipeline stall.
            return (uint)((codepoint | 32) - 97) <= 25U;
        }

        // Is codepoint a character per XML 1.0 spec (2.2)?
        static bool IsXml10Char(int codepoint)
        {
            // Fast path: printable ASCII, the overwhelming majority of characters inside a text
            // literal, is always valid per XML 1.0's Char production - verified to agree with
            // XmlConvert.IsXmlChar for the entire 0x00-0x7F range (only 0x00-0x08, 0x0B, 0x0C and
            // 0x0E-0x1F are invalid there, none of them in this range). Skips the property lookup
            // table XmlConvert.IsXmlChar uses internally for the common case.
            if (codepoint is >= 0x20 and <= 0x7F)
            {
                return true;
            }

            // .NET is using a lookup table with properties
            if (codepoint <= 0xFFFF)
            {
                return XmlConvert.IsXmlChar((char)codepoint);
            }

            return codepoint <= 0x10FFFF;
        }
    }

    /// <summary>
    /// R1C1 mode only, called with <see cref="_c"/> at a leading <c>R</c>/<c>C</c> (either case):
    /// attempt to scan a full R1C1 reference token - a row (<c>R</c>, <c>R5</c>, <c>R[-14]</c>), a
    /// column (<c>C</c>, <c>C5</c>, <c>C[-14]</c>), or a cell (row immediately followed by column,
    /// e.g. <c>RC</c>, <c>R1C1</c>, <c>R[-1]C[-1]</c>). Unlike A1, where a cell's row/column are
    /// both plain digit runs already covered by ordinary identifier scanning, R1C1's bracketed
    /// relative form (<c>[-14]</c>) is not itself ident-continuation text, so it needs its own scan
    /// to end up as a single token rather than three (<c>Ident "R"</c>, <c>SquareIdent "[-14]"</c>,
    /// ...). Purely speculative: nothing is consumed on failure (a token whose text isn't fully
    /// R1C1-shaped, or one that's followed by further identifier characters - e.g. "R1C1style",
    /// where the longer NAME must win, matching the oracle's own maximal-munch lexer), so the caller
    /// can safely fall back to ordinary identifier scanning from the very first character.
    /// </summary>
    private Token? TryScanR1C1Reference()
    {
        int i = this._i;
        bool matchedRow = TryScanR1C1Axis(this._input, ref i, 'R');
        bool matchedColumn = TryScanR1C1Axis(this._input, ref i, 'C');
        if (!matchedRow && !matchedColumn)
        {
            return null;
        }

        if (i < this._input.Length && IsR1C1IdentNext(this._input[i]))
        {
            return null;
        }

        this._i = i - 1;
        this.Advance();
        return this.T(TokenType.Ident);

        // Duplicates IsIdentStart/IsIdentNext from Next() (both local functions, out of reach from
        // here) - true for any character that could extend a NAME beyond this R1C1-shaped run.
        static bool IsR1C1IdentNext(char c)
        {
            return
                (uint)((c | 0x20) - 'a') <= 25u || // ASCII letter
                c is '$' or '_' or '\\' or '?' or '.' ||
                c is >= '0' and <= '9' ||
                c > 0x7F;
        }
    }

    /// <summary>
    /// Try to scan one axis (<paramref name="axisLetter"/> is <c>'R'</c> or <c>'C'</c>, matched case
    /// insensitively) starting at <c>input[i]</c>, advancing <paramref name="i"/> past it on
    /// success. Grammar: the axis letter, then either <c>"[" ("-")? digit+ "]"</c> (relative, e.g.
    /// <c>R[-14]</c>), digit+ (absolute, e.g. <c>R14</c>), or nothing at all (relative zero, e.g.
    /// bare <c>R</c>). Range validation of the digits themselves (row/column bounds) is left to the
    /// actual decode in <see cref="TokenParser.ParseR1C1Reference(ReadOnlySpan{char},ref int)"/> -
    /// this only needs to find the token's extent. Returns false, leaving <paramref name="i"/>
    /// untouched, if the letter itself doesn't match or an opened bracket is never closed over a
    /// digit run.
    /// </summary>
    private static bool TryScanR1C1Axis(string input, ref int i, char axisLetter)
    {
        if (i >= input.Length || (input[i] | 0x20) != (axisLetter | 0x20))
        {
            return false;
        }

        int j = i + 1;
        if (j < input.Length && input[j] == '[')
        {
            j++;
            if (j < input.Length && input[j] == '-')
            {
                j++;
            }

            int digitsStart = j;
            while (j < input.Length && input[j] is >= '0' and <= '9')
            {
                j++;
            }

            if (j == digitsStart || j >= input.Length || input[j] != ']')
            {
                return false;
            }

            i = j + 1;
            return true;
        }

        while (j < input.Length && input[j] is >= '0' and <= '9')
        {
            j++;
        }

        // Zero digits consumed here is fine - bare "R"/"C" is a valid relative-zero shorthand.
        i = j;
        return true;
    }

    private Token T(TokenType type)
    {
        return new Token(type, this._start, this._i);
    }

    private int Advance()
    {
        if (++this._i >= this._input.Length)
        {
            this._c = -1;
            return (char)this._c;
        }

        char c = this._input[this._i];

        // Fast path: formulas are overwhelmingly ASCII, and no ASCII codepoint is ever a
        // surrogate, so skip both surrogate checks below entirely for the common case.
        if (c < 0x80)
        {
            return this._c = c;
        }

        if (char.IsLowSurrogate(c))
        {
            throw ParsingException.UnpairedSurrogate(c, this._i);
        }

        if (char.IsHighSurrogate(c))
        {
            if (this._i + 1 >= this._input.Length)
            {
                throw ParsingException.UnpairedSurrogate(c, this._i);
            }

            char low = this._input[++this._i];
            if (!char.IsLowSurrogate(low))
            {
                throw ParsingException.UnpairedSurrogate(c, this._i - 1);
            }

            this._c = char.ConvertToUtf32(c, low);
            return this._c;
        }

        return this._c = c;
    }
}

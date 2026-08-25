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

    public Lexer(string input)
    {
        this.Reset(input);
    }

    private bool IsEof => this._c == EOF;

    /// <summary>
    /// Prepare lexer to start tokenization of the <paramref name="formula"/>.
    /// </summary>
    /// <param name="formula">Formula to tokenize.</param>
    public void Reset(string formula)
    {
        this._input = formula ?? throw new ArgumentNullException();
        this._start = -1;
        this._i = -1;
        this._c = 0;
        this._lookaheadHead = 0;
        this._lookaheadCount = 0;
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

                        if (char2 == 'A')
                        {
                            return Error("#NAME?", 3);
                        }

                        int char3 = this.Advance();
                        if (char2 == 'U' && char3 == 'L')
                        {
                            return Error("#NULL!", 4);
                        }

                        if (char2 == 'U' && char3 == 'M')
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

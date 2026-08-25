namespace XlsxSharp.Parser.Pratt;

/// <summary>
/// Pratt parser.
/// </summary>
internal class Parser<T, TContext>
{
    // TokenType is a small, dense, zero-based enum - an array indexed by (int)TokenType is a
    // direct lookup instead of a Dictionary<TokenType,_> hash, on the hottest dispatch path of the
    // parser (every atom goes through ParseAtom, every operator through ParseExpression's loop).
    private static readonly int TokenTypeCount = Enum.GetValues<TokenType>().Length;

    private readonly Lexer _lexer = new();
    private readonly IPrefixParselet<T, TContext>?[] _prefixParselets = new IPrefixParselet<T, TContext>?[TokenTypeCount];
    private readonly IParselet<T, TContext>?[] _parselets = new IParselet<T, TContext>?[TokenTypeCount];

    internal string Input { get; private set; } = string.Empty;

    /// <summary>
    /// True while parsing an R1C1-style formula rather than A1, set for the duration of the current
    /// <see cref="ParseFormula"/> call. Consulted by <c>ParserExtensions</c>' reference-shape checks
    /// so that every parselet (registered once, shared between both styles) stays style-agnostic.
    /// </summary>
    internal bool IsR1C1 { get; private set; }

    /// <summary>
    /// Combines two pure-reference operands (see <see cref="Node{T}.IsPureReference"/>) of the
    /// range operator (<c>:</c>) into a single node, set once by <see cref="ParserFactory"/>. This
    /// lives behind a delegate rather than a direct <c>IAstFactory{...}.BinaryNode</c> call because
    /// <see cref="Parser{T,TContext}"/> is deliberately not generic over <c>TScalarValue</c>.
    /// </summary>
    internal Func<TContext, SymbolRange, T, T, T>? RangeCombiner { private get; set; }

    /// <summary>
    /// Combines two pure-reference operands of the union operator (<c>,</c> - not the argument
    /// separator or array-row separator) into a single node, set once by <see cref="ParserFactory"/>.
    /// See the note on <see cref="RangeCombiner"/> for why this is a delegate.
    /// </summary>
    internal Func<TContext, SymbolRange, T, T, T>? UnionCombiner { private get; set; }

    /// <summary>
    /// Combines two pure-reference operands of the reference intersection operator (whitespace
    /// between two references, e.g. <c>NamedRange1 NamedRange2</c>) into a single node, set once
    /// by <see cref="ParserFactory"/>. See the note on <see cref="RangeCombiner"/> for why this is
    /// a delegate.
    /// </summary>
    internal Func<TContext, SymbolRange, T, T, T>? IntersectionCombiner { private get; set; }

    /// <summary>
    /// True while parsing a function-call argument directly (see
    /// <c>ParserExtensions.ParseArgumentList</c>): there, a top-level <c>,</c> means "next
    /// argument", not "union" - <c>SUM(A1,B2)</c> is two arguments, not one unioned argument. A
    /// fresh pair of parentheses (see <see cref="Parselets.GroupParselet{T,TContext}"/>) always
    /// resets this to <c>false</c> for its own content, since <c>(A1,B2)</c> is a union even as a
    /// single argument, e.g. in <c>SUM((A1,B2))</c> - matching the oracle's own
    /// <c>skipRangeUnion</c> parameter, threaded the same way.
    /// </summary>
    internal bool SkipUnion { get; set; }

    public T ParseFormula(string formula, TContext ctx, bool isR1C1 = false)
    {
        this.Input = formula;
        this.IsR1C1 = isR1C1;
        this._lexer.Reset(formula, isR1C1);
        Node<T> node = this.ParseExpression(ctx, 0);

        // Trailing whitespace is insignificant at the end of a formula (mirrors the TrimEnd()
        // done before FormulaParser<TScalarValue,TNode,TContext> tokenizes with the Rolex lexer).
        Token next = this._lexer.Peek();
        if (next.Type == TokenType.Whitespace)
        {
            this._lexer.Consume();
            next = this._lexer.Peek();
        }

        if (next.Type != TokenType.Eof)
        {
            throw ParsingException.TrailingToken(next.Range.Start, next.Type);
        }

        return node.Value;
    }

    internal Node<T> ParseExpression(TContext ctx, int minBp)
    {
        Node<T> node = this.Prefix(ctx);

        while (true)
        {
            // Whitespace is insignificant everywhere an operator is expected (e.g. "1 + 2",
            // "SUM(A1, B2)"). It is *not* insignificant everywhere in the full grammar (a run of
            // whitespace between two references is the reference intersection operator), but that
            // operator isn't implemented, so a real intersection still correctly fails below: the
            // token after the (now-consumed) whitespace won't be a registered operator either, and
            // whatever's left over is rejected as a trailing token by the caller.
            this.SkipWhitespace();

            Token maybeOp = this._lexer.Peek();
            if (maybeOp.Type == TokenType.Eof)
            {
                break;
            }

            IParselet<T, TContext>? parselet = this._parselets[(int)maybeOp.Type];
            if (parselet is null)
            {
                break;
            }

            int bp = parselet!.GetBindingPower();
            if (bp <= minBp)
            {
                break;
            }

            Token op = this._lexer.Consume();
            node = parselet.Parse(ctx, node, op);
        }

        return node;
    }

    private Node<T> Prefix(TContext ctx)
    {
        Node<T> node = this.ParseIntersectedAtom(ctx);
        return this.ParseUnionChain(ctx, node);
    }

    /// <summary>
    /// A reference atom (see <see cref="ParseReferenceAtom"/>) plus any intersection chain
    /// immediately on it (see <see cref="ParseIntersectionChain"/>) - i.e. everything up to, but
    /// not including, a union. Used both by <see cref="Prefix"/> and directly as each operand of a
    /// union: like a range operand, a union operand is never itself a further union - "A1,B2,C3"
    /// is a left-associative chain of two union operations, not one union whose right side is
    /// itself a union.
    /// </summary>
    private Node<T> ParseIntersectedAtom(TContext ctx)
    {
        Node<T> node = this.ParseReferenceAtom(ctx);
        return this.ParseIntersectionChain(ctx, node);
    }

    /// <summary>
    /// A single atom (see <see cref="ParseAtom"/>) plus any range chain immediately on it (see
    /// <see cref="ParseRangeChain"/>) - i.e. everything up to, but not including, an intersection
    /// or a union. Used both by <see cref="ParseIntersectedAtom"/> and directly as each operand of
    /// an intersection - unlike a range operand (which is a single bare atom, see
    /// <see cref="ParseRangeChain"/>), an intersection operand can itself carry a range, e.g.
    /// "A1 A2:B2" is Intersection(A1, Range(A2,B2)).
    /// </summary>
    private Node<T> ParseReferenceAtom(TContext ctx)
    {
        Node<T> node = this.ParseAtom(ctx);
        return this.ParseRangeChain(ctx, node);
    }

    /// <summary>
    /// Dispatch a single prefix parselet for the upcoming token, without any range- or union-
    /// chaining. This is also used directly (not through <see cref="Prefix"/>) to parse the
    /// right-hand operand of a range: that operand must itself be a single atom, never a range -
    /// <c>A1:B2:C3</c> is a left-associative chain of two range operations, not one range whose
    /// right side is itself a range.
    /// </summary>
    private Node<T> ParseAtom(TContext ctx)
    {
        this.SkipWhitespace();
        Token token = this._lexer.Consume();

        IPrefixParselet<T, TContext>? parselet = this._prefixParselets[(int)token.Type];
        if (parselet is null)
        {
            throw new InvalidOperationException($"No parselet found for {token.Type}.");
        }

        return parselet.Parse(ctx, token);
    }

    /// <summary>
    /// After a pure-reference atom (see <see cref="Node{T}.IsPureReference"/>), consume any
    /// following <c>:</c> range operator(s) left-associatively: <c>A1:B2:C3</c> parses as
    /// <c>(A1:B2):C3</c>, matching the oracle's own left-folding loop. The range operator binds
    /// tighter than every other operator (even unary +/- and %, which recurse into
    /// <see cref="ParseExpression"/> - i.e. back into <see cref="Prefix"/> - for their own operand,
    /// so a chain here is already fully resolved by the time they see it), which is why this lives
    /// inside <see cref="Prefix"/> rather than as a normal <see cref="IParselet{T,TContext}"/>
    /// registered at some binding power.
    /// </summary>
    private Node<T> ParseRangeChain(TContext ctx, Node<T> left)
    {
        if (!left.IsPureReference)
        {
            return left;
        }

        while (true)
        {
            // The oracle's COLON lexer token absorbs surrounding whitespace directly into its own
            // regex, so "A1 : B2" is indistinguishable from "A1:B2" to it - peek past (at most) one
            // insignificant whitespace token to see whether a ":" is really there before consuming
            // anything. Unlike a plain unconditional skip, this leaves the whitespace untouched
            // when it's not followed by ":", so a real reference intersection (a run of whitespace
            // between two references with no ":" in it) is still visible to
            // <see cref="ParseIntersectionChain"/> afterward.
            Token maybeColon = this.PeekPastWhitespace(out bool hadWhitespace);
            if (maybeColon.Type != TokenType.Range)
            {
                break;
            }

            if (hadWhitespace)
            {
                this._lexer.Consume();
            }

            this._lexer.Consume();
            this.SkipWhitespace();
            Node<T> right = this.ParseAtom(ctx);
            if (!right.IsPureReference)
            {
                throw new ParsingException($"Unable to parse value starting from position {right.Range.Start}.");
            }

            SymbolRange range = new(left.Range.Start, right.Range.End);
            T value = this.RangeCombiner!(ctx, range, left.Value, right.Value);
            left = new Node<T>(value, range, isPureReference: true);
        }

        return left;
    }

    /// <summary>
    /// After a pure-reference atom-plus-range-chain (see <see cref="ParseReferenceAtom"/>), consume
    /// any following reference intersection(s) left-associatively - a run of whitespace directly
    /// between two references, with nothing else in it. This mirrors the oracle exactly at the
    /// lexer level rather than needing any backtracking: every one of its operator/paren/comma
    /// tokens absorbs surrounding whitespace directly into its own regex (see e.g. PLUS or COMMA in
    /// LexerA1.rl), so its lexer only ever emits a literal SPACE token between two tokens that
    /// aren't otherwise claimed by anything - i.e. exactly the reference-intersection case. Once
    /// that's confirmed here (a whitespace token immediately followed by one of the "content"
    /// token types - the ones with no such absorption, listed in <see cref="IsContentTokenType"/>),
    /// there's nothing to backtrack for either: if the right side then turns out not to be
    /// reference-shaped, the oracle doesn't fall back to treating the whitespace as insignificant,
    /// it just fails the whole parse - so this does too.
    /// </summary>
    private Node<T> ParseIntersectionChain(TContext ctx, Node<T> left)
    {
        if (!left.IsPureReference)
        {
            return left;
        }

        while (true)
        {
            Token candidate = this.PeekPastWhitespace(out bool hadWhitespace);
            if (!hadWhitespace || !IsContentTokenType(candidate.Type))
            {
                break;
            }

            this._lexer.Consume();
            Node<T> right = this.ParseReferenceAtom(ctx);
            if (!right.IsPureReference)
            {
                throw new ParsingException($"Unable to parse value starting from position {right.Range.Start}.");
            }

            SymbolRange range = new(left.Range.Start, right.Range.End);
            T value = this.IntersectionCombiner!(ctx, range, left.Value, right.Value);
            left = new Node<T>(value, range, isPureReference: true);
        }

        return left;
    }

    /// <summary>
    /// Peek the token that would follow an optional single insignificant whitespace token, without
    /// consuming anything.
    /// </summary>
    private Token PeekPastWhitespace(out bool hadWhitespace)
    {
        Token first = this._lexer.Peek(1);
        if (first.Type != TokenType.Whitespace)
        {
            hadWhitespace = false;
            return first;
        }

        hadWhitespace = true;
        return this._lexer.Peek(2);
    }

    /// <summary>
    /// The token types whose oracle-side lexer counterpart doesn't absorb surrounding whitespace
    /// into its own token (unlike every operator, paren, comma, and semicolon) - i.e. the ones that
    /// can be preceded by a genuine, separately-tokenized whitespace run. Only these can start the
    /// right-hand operand of a reference intersection - see <see cref="ParseIntersectionChain"/>.
    /// </summary>
    private static bool IsContentTokenType(TokenType type)
    {
        return type
            is TokenType.Ident
                or TokenType.Number
                or TokenType.Text
                or TokenType.Error
                or TokenType.QIdent
                or TokenType.SquareIdent;
    }

    /// <summary>
    /// After a pure-reference atom-plus-range-plus-intersection-chain (see
    /// <see cref="ParseIntersectedAtom"/>), and unless <see cref="SkipUnion"/> says a "," here
    /// means something else (an argument separator), consume any following <c>,</c> union
    /// operator(s) left-associatively - matching the oracle's own left-folding loop, the outermost
    /// layer of its reference sub-grammar (looser-binding than range or intersection, which is why
    /// this wraps <see cref="ParseIntersectedAtom"/> rather than being folded into
    /// <see cref="ParseRangeChain"/> or <see cref="ParseIntersectionChain"/>). Still binds tighter
    /// than every arithmetic operator, for the same reason range does - see the note there. Unlike
    /// range/intersection, whitespace around "," needs no special handling here: COMMA absorbs it
    /// into its own token on the oracle side exactly like every other operator, so an unconditional
    /// skip (matching how every other operator in this parser is looked for) is already correct.
    /// </summary>
    private Node<T> ParseUnionChain(TContext ctx, Node<T> left)
    {
        if (!left.IsPureReference || this.SkipUnion)
        {
            return left;
        }

        while (true)
        {
            this.SkipWhitespace();
            if (this._lexer.Peek().Type != TokenType.Comma)
            {
                break;
            }

            this._lexer.Consume();
            this.SkipWhitespace();
            Node<T> right = this.ParseIntersectedAtom(ctx);
            if (!right.IsPureReference)
            {
                throw new ParsingException($"Unable to parse value starting from position {right.Range.Start}.");
            }

            SymbolRange range = new(left.Range.Start, right.Range.End);
            T value = this.UnionCombiner!(ctx, range, left.Value, right.Value);
            left = new Node<T>(value, range, isPureReference: true);
        }

        return left;
    }

    internal void SkipWhitespace()
    {
        if (this._lexer.Peek().Type == TokenType.Whitespace)
        {
            this._lexer.Consume();
        }
    }

    public Token LookAhead(int distance)
    {
        return this._lexer.Peek(distance);
    }

    internal Token Consume(TokenType expectedType)
    {
        Token token = this._lexer.Consume();
        if (token.Type != expectedType)
        {
            throw new InvalidOperationException($"Expected token of type {expectedType}, but received {token.Type}.");
        }

        return token;
    }

    internal Token Consume()
    {
        return this._lexer.Consume();
    }

    internal void Register(TokenType type, IPrefixParselet<T, TContext> parselet)
    {
        if (this._prefixParselets[(int)type] is not null)
        {
            throw new ArgumentException($"A prefix parselet is already registered for {type}.", nameof(type));
        }

        this._prefixParselets[(int)type] = parselet;
    }

    internal void Register(TokenType type, IParselet<T, TContext> parselet)
    {
        if (this._parselets[(int)type] is not null)
        {
            throw new ArgumentException($"A parselet is already registered for {type}.", nameof(type));
        }

        this._parselets[(int)type] = parselet;
    }
}

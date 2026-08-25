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
    /// Wraps a pure-reference operand of the postfix spill-range operator (<c>#</c>, e.g.
    /// <c>A1#</c>) into a single node, set once by <see cref="ParserFactory"/>. See the note on
    /// <see cref="RangeCombiner"/> for why this is a delegate.
    /// </summary>
    internal Func<TContext, SymbolRange, T, T>? SpillCombiner { private get; set; }

    /// <summary>
    /// Wraps a node whose given range extends past what its own value's range covers (currently
    /// only leading whitespace before the whole formula - see <see cref="ParseFormula"/> - but
    /// this is the same "extra raw text surrounds an inner node" shape <see cref="Parselets.GroupParselet{TScalar,T,TContext}"/>
    /// uses for parens), set once by <see cref="ParserFactory"/>. See the note on
    /// <see cref="RangeCombiner"/> for why this is a delegate.
    /// </summary>
    internal Func<TContext, SymbolRange, T, T>? NestedCombiner { private get; set; }

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

        // SkipUnion is saved/restored in pairs (see ParseArgumentList, GroupParselet), but a
        // parselet throwing between the save and its matching restore - e.g. a malformed argument
        // like "SUM(A1:" - skips the restore. A single Parser instance is routinely reused across
        // many formulas (every production call site does this), so that leftover `true` would
        // silently corrupt every top-level comma in every formula parsed afterwards, turning a
        // union into a rejected trailing token. Force it back to the real default here so no
        // exception from a previous, unrelated formula can ever survive into this one.
        this.SkipUnion = false;

        this._lexer.Reset(formula, isR1C1);

        // Leading whitespace is significant to the oracle only through which token follows it:
        // every operator/paren/brace/comma/etc. token absorbs surrounding whitespace directly
        // into its own lexer regex (see IsContentTokenType), so e.g. "  -8" and "-8" are the same
        // token stream to it, and the MINUS token's own range already starts at index 0 either
        // way. A "content" token (Ident, Number, Text, Error, QIdent, SquareIdent) never absorbs
        // it, so leading whitespace before one of those is genuinely dropped, not carried by any
        // token's range. ParseAtom's own SkipWhitespace() always discards it either way, so the
        // only way to preserve it here (needed by a text-reconstructing consumer such as
        // FormulaConverter, via IAstFactory.Nested) is to detect the case up front and wrap the
        // whole result in a Nested-shaped span starting at 0.
        Token firstToken = this.PeekPastWhitespace(out bool hasLeadingWhitespace);
        bool wrapLeadingWhitespace = hasLeadingWhitespace && !IsContentTokenType(firstToken.Type);

        Node<T> node = this.ParseExpression(ctx, 0);

        // Trailing whitespace is insignificant at the end of a formula (mirrors the TrimEnd() the
        // removed recursive-descent parser did before tokenizing with the Rolex lexer).
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

        T value = wrapLeadingWhitespace
            ? this.NestedCombiner!(ctx, new SymbolRange(0, node.Range.End), node.Value)
            : node.Value;
        return value;
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
        node = this.ParseSpillChain(ctx, node);
        return this.ParseRangeChain(ctx, node);
    }

    /// <summary>
    /// After a pure-reference atom, consume a following postfix spill-range operator (<c>#</c>,
    /// e.g. <c>A1#</c>), if present. Lives here - between the bare atom and any range chain -
    /// rather than as a normal <see cref="IParselet{T,TContext}"/> registered at some binding
    /// power, for the same reason <see cref="ParseRangeChain"/> does: <c>F2#:A7</c> must be
    /// <c>Range(SpillRange(F2), A7)</c>, so the spill has to apply *before* the range chain sees
    /// its left operand, not after the whole <see cref="Prefix"/> chain completes. The result
    /// stays a pure reference (a spill range is a valid range/union/intersection operand, e.g.
    /// <c>A1#:B5</c> or <c>A1#,B1</c>), unlike a "value" unary operator such as <c>%</c>.
    /// </summary>
    private Node<T> ParseSpillChain(TContext ctx, Node<T> left)
    {
        if (!left.IsPureReference)
        {
            return left;
        }

        Token maybeSpill = this.PeekPastWhitespace(out bool hadWhitespace);
        if (maybeSpill.Type != TokenType.Spill)
        {
            return left;
        }

        if (hadWhitespace)
        {
            this._lexer.Consume();
        }

        Token spillToken = this._lexer.Consume();
        SymbolRange range = new(left.Range.Start, spillToken.Range.End);
        T value = this.SpillCombiner!(ctx, range, left.Value);
        return new Node<T>(value, range, isPureReference: true);
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
            right = this.ParseSpillChain(ctx, right);
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
            // Unlike ParseAtom's "No parselet found for X" (a token type with no grammar coverage
            // at all - a genuine gap), every call site here has already committed to a specific
            // grammar rule (an argument list, a range, a sheet prefix, ...) and is only checking
            // that rule's next required token is actually present. A mismatch is always a
            // malformed formula (e.g. "SUM(A1" with no closing paren), not a missing feature, so
            // this is a ParsingException like every other formula-rejection reason.
            throw new ParsingException($"Expected token of type {expectedType}, but received {token.Type}.");
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

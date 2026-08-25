namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Get an error node from a <see cref="TokenType.Error"/> token, e.g. <c>#REF!</c> or <c>#N/A</c>.
/// <see cref="IAstFactory{TScalarValue,TNode,TContext}.ErrorNode"/> expects upper case text no
/// matter the casing used in the formula (e.g. <c>#div/0!</c> is still <c>#DIV/0!</c>).
/// </summary>
internal class ErrorParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public ErrorParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        string error = token.GetText(this._parser.Input).ToString().ToUpperInvariant();

        // Only "#REF!" is reference-shaped (a deleted reference) and so a valid range operand -
        // every other error (#N/A, #VALUE!, ...) is a plain value, matching the oracle's lexer,
        // which tokenizes "#REF!" alone as REF_CONSTANT and everything else as NONREF_ERRORS.
        bool isPureReference = string.Equals(error, "#REF!", StringComparison.Ordinal);
        SymbolRange range = isPureReference ? this.ConsumeSwallowedReference(token.Range) : token.Range;

        T node = this._factory.ErrorNode(ctx, range, error);
        return new Node<T>(node, range, isPureReference);
    }

    /// <summary>
    /// A bare "#REF!" directly (no whitespace) followed by another "#REF!", a bare A1 cell
    /// (optionally extended to an area via ":cell2"), or a row/column span still just means
    /// "#REF!" - the oracle's own REF_CONSTANT ref-atom production parses and discards the
    /// trailing reference rather than rejecting it or combining it into anything, presumably to
    /// tolerate this specific shape of corrupted/legacy formula text. With whitespace in between
    /// it's not this construct at all, but a (not yet implemented) reference intersection instead -
    /// e.g. "#REF! A1" is Intersection(#REF!, A1) to the oracle, not this swallowing.
    /// </summary>
    private SymbolRange ConsumeSwallowedReference(SymbolRange errorRange)
    {
        Token next = this._parser.LookAhead(1);
        if (next.Type == TokenType.Error && ParserExtensions.EqualCaseInsensitive(next.GetText(this._parser.Input), "#REF!"))
        {
            Token consumed = this._parser.Consume(TokenType.Error);
            return new SymbolRange(errorRange.Start, consumed.Range.End);
        }

        if (next.Type is TokenType.Ident or TokenType.Number)
        {
            ReadOnlySpan<char> text = next.GetText(this._parser.Input);
            bool looksLikeReference =
                ParserExtensions.TryGetCellA1(text, out _)
                || ParserExtensions.TryGetColA1(text, out _)
                || ParserExtensions.TryGetRowA1(text, out _);
            if (looksLikeReference)
            {
                Token consumed = this._parser.Consume();
                this._parser.TryReferenceA1(consumed, out _, out SymbolRange consumedRange);
                return new SymbolRange(errorRange.Start, consumedRange.End);
            }
        }

        return errorRange;
    }
}

namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// A bang reference (<see cref="TokenType.Bang"/> at the start of an atom, e.g. <c>!$B1</c> or
/// <c>!A1:B2</c>) - a leftover reference to a name that's been deleted, keeping only the cell(s) it
/// used to point at. Unlike every other use of <c>!</c> in this parser (a sheet qualifier), this
/// one requires the reference to immediately follow with no whitespace at all: the oracle's own
/// BANG_REFERENCE lexer token fuses the "!" and the reference into one token, so "! A1" doesn't
/// even tokenize on the oracle side, let alone parse.
/// </summary>
internal class BangReferenceParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public BangReferenceParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token bangToken)
    {
        Token next = this._parser.LookAhead(1);

        // "!#REF!" (any casing) collapses to a plain #REF! error, not a BangReference - and unlike
        // every other "!...#REF!" collapse in this parser (e.g. sheet!#REF!), the oracle normalizes
        // the casing here (it passes a fixed "#REF!" literal, not the raw matched text).
        if (next.Type == TokenType.Error && ParserExtensions.EqualCaseInsensitive(next.GetText(this._parser.Input), "#REF!"))
        {
            Token consumed = this._parser.Consume(TokenType.Error);
            SymbolRange range = new(bangToken.Range.Start, consumed.Range.End);
            T value = this._factory.ErrorNode(ctx, range, "#REF!");
            return new Node<T>(value, range, isPureReference: true);
        }

        if (next.Type is TokenType.Ident or TokenType.Number)
        {
            ReadOnlySpan<char> text = next.GetText(this._parser.Input);
            if (this._parser.IsAnyReferenceShape(text))
            {
                Token consumed = this._parser.Consume();
                this._parser.TryReferenceA1(consumed, out ReferenceArea area, out SymbolRange areaRange);
                SymbolRange range = new(bangToken.Range.Start, areaRange.End);
                T value = this._factory.BangReference(ctx, range, area);
                return new Node<T>(value, range, isPureReference: true);
            }
        }

        throw new ParsingException($"Unable to parse value starting from position {bangToken.Range.Start}.");
    }
}

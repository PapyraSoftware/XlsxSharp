namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// A quoted sheet name/sheet range prefix (<see cref="TokenType.QIdent"/>), e.g. <c>'New York'!A1</c>,
/// <c>'Jane''s'!name</c> or a quoted 3D reference <c>'Sheet 1:Sheet 2'!A1</c> (the colon separating
/// the two sheet names is inside the quotes here, unlike the unquoted 3D reference case handled by
/// <see cref="IdentParselet{TScalar,T,TContext}"/>).
/// </summary>
/// <remarks>
/// A quoted external workbook reference (e.g. <c>'[2]Sheet 1'!A1</c>) isn't implemented yet - that
/// needs the square-bracket external reference syntax, which isn't recognized anywhere else in the
/// pratt parser either.
/// </remarks>
internal class QIdentParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public QIdentParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        if (this._parser.LookAhead(1).Type != TokenType.Bang)
        {
            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        string unquoted = Unquote(token.GetText(this._parser.Input));
        if (unquoted.StartsWith('['))
        {
            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        Token bangToken = this._parser.Consume(TokenType.Bang);
        SymbolRange withBangRange = token.Range.ExtendRight(bangToken.Range);
        Token refToken = this._parser.Consume();

        int colonIndex = unquoted.IndexOf(':');
        if (colonIndex < 0)
        {
            if (this._parser.TryReferenceA1(refToken, out ReferenceArea area, out SymbolRange areaRange))
            {
                SymbolRange range = withBangRange.ExtendRight(areaRange);
                T value = this._factory.SheetReference(ctx, range, unquoted, area);
                return new Node<T>(value, range);
            }

            if (this._parser.TryGetName(refToken, out ReadOnlySpan<char> name))
            {
                SymbolRange range = withBangRange.ExtendRight(refToken.Range);
                T value = this._factory.SheetName(ctx, range, unquoted, name.ToString());
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        string firstSheet = unquoted[..colonIndex];
        string secondSheet = unquoted[(colonIndex + 1)..];
        if (this._parser.TryReferenceA1(refToken, out ReferenceArea range3DArea, out SymbolRange range3DAreaRange))
        {
            SymbolRange range = withBangRange.ExtendRight(range3DAreaRange);
            T value = this._factory.Reference3D(ctx, range, firstSheet, secondSheet, range3DArea);
            return new Node<T>(value, range);
        }

        throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
    }

    /// <summary>
    /// Strip the surrounding single quotes and collapse escaped <c>''</c> pairs into a single
    /// <c>'</c>, e.g. <c>'Jane''s'</c> becomes <c>Jane's</c>.
    /// </summary>
    private static string Unquote(ReadOnlySpan<char> quotedText)
    {
        ReadOnlySpan<char> inner = quotedText[1..^1];
        if (inner.IndexOf('\'') < 0)
        {
            return inner.ToString();
        }

        Span<char> buffer = new char[inner.Length];
        int w = 0;
        int i = 0;
        while (i < inner.Length)
        {
            if (inner[i] == '\'')
            {
                i++;
            }

            buffer[w++] = inner[i++];
        }

        return buffer[..w].ToString();
    }
}

namespace XlsxSharp.Parser.Pratt;

/// <summary>
/// Pratt parser.
/// </summary>
internal class Parser<T, TContext>
{
    private readonly Lexer _lexer = new();
    private readonly Dictionary<TokenType, IPrefixParselet<T, TContext>> _prefixParselets = new();
    private readonly Dictionary<TokenType, IParselet<T, TContext>> _parselets = new();

    internal string Input { get; private set; } = string.Empty;

    public T ParseFormula(string formula, TContext ctx)
    {
        this.Input = formula;
        this._lexer.Reset(formula);
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

            bool isOp = this._parselets.TryGetValue(maybeOp.Type, out IParselet<T, TContext>? parselet);
            if (!isOp)
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
        this.SkipWhitespace();
        Token token = this._lexer.Consume();

        if (!this._prefixParselets.TryGetValue(token.Type, out IPrefixParselet<T, TContext>? parselet))
        {
            throw new InvalidOperationException($"No parselet found for {token.Type}.");
        }

        return parselet.Parse(ctx, token);
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
        this._prefixParselets.Add(type, parselet);
    }

    internal void Register(TokenType type, IParselet<T, TContext> parselet)
    {
        this._parselets.Add(type, parselet);
    }
}

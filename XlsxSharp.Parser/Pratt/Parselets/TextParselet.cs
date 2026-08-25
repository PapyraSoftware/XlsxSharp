namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Get a text node from a <see cref="TokenType.Text"/> token, e.g. <c>"abc"</c> or <c>"a""b"</c>
/// (an escaped quote, unescaping to <c>a"b</c>).
/// </summary>
internal class TextParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public TextParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        string text = ParserExtensions.UnescapeText(token.GetText(this._parser.Input));
        T node = this._factory.TextNode(ctx, token.Range, text);
        return new Node<T>(node, token.Range);
    }
}

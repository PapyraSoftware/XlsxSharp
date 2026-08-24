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
        T node = this._factory.ErrorNode(ctx, token.Range, error);
        return new Node<T>(node, token.Range);
    }
}

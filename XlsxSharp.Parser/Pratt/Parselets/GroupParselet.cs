namespace XlsxSharp.Parser.Pratt.Parselets;

internal class GroupParselet<T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly Parser<T, TContext> _parser;

    public GroupParselet(Parser<T, TContext> parser)
    {
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token leftParen)
    {
        Node<T> node = this._parser.ParseExpression(ctx, 0);
        Token rightParen = this._parser.Consume(TokenType.RightParen);

        // Not node.ExtendLeft(leftParen).ExtendRight(rightParen): that asserts strict adjacency
        // between the parens and the inner expression, which no longer holds now that whitespace
        // is allowed there (e.g. "( 1 + 2 )").
        return new Node<T>(node.Value, new SymbolRange(leftParen.Range.Start, rightParen.Range.End));
    }
}

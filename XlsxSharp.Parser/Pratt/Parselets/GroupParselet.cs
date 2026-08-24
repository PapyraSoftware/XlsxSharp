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
        return node.ExtendLeft(leftParen).ExtendRight(rightParen);
    }
}

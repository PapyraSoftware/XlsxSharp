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
        // A fresh "(...)" always allows union for its own content, regardless of the outer
        // context - e.g. "SUM((A1,B2))" is one argument containing a union, even though a bare
        // "," directly inside SUM(...) would otherwise mean "next argument" (see Parser.SkipUnion).
        bool previousSkipUnion = this._parser.SkipUnion;
        this._parser.SkipUnion = false;
        Node<T> node = this._parser.ParseExpression(ctx, 0);
        this._parser.SkipUnion = previousSkipUnion;
        Token rightParen = this._parser.Consume(TokenType.RightParen);

        // Not node.ExtendLeft(leftParen).ExtendRight(rightParen): that asserts strict adjacency
        // between the parens and the inner expression, which no longer holds now that whitespace
        // is allowed there (e.g. "( 1 + 2 )").
        //
        // IsPureReference propagates from the inner expression: matches the oracle's "ambiguity"
        // backtracking, where a parenthesized ref-expression (e.g. "(A1:B2)") remains a valid
        // operand of an enclosing range/intersection/spill operator - "(A1:B2):C3" is accepted.
        return new Node<T>(node.Value, new SymbolRange(leftParen.Range.Start, rightParen.Range.End), node.IsPureReference);
    }
}

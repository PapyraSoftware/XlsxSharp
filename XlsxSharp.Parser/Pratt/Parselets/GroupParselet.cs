namespace XlsxSharp.Parser.Pratt.Parselets;

internal class GroupParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public GroupParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
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
        SymbolRange range = new(leftParen.Range.Start, rightParen.Range.End);

        // factory.Nested(...), not node.Value directly: a value-only consumer (e.g. the calc
        // engine's own AstFactory) can treat it as a no-op passthrough - parens don't change what
        // an expression evaluates to - but a text-reconstructing consumer (CopyVisitor, used by
        // FormulaConverter) needs this call to know the parens themselves must be preserved in the
        // output text, which it can't infer from the inner node's value alone.
        T value = this._factory.Nested(ctx, range, node.Value);

        // IsPureReference propagates from the inner expression: matches the oracle's "ambiguity"
        // backtracking, where a parenthesized ref-expression (e.g. "(A1:B2)") remains a valid
        // operand of an enclosing range/intersection/spill operator - "(A1:B2):C3" is accepted.
        return new Node<T>(value, range, node.IsPureReference);
    }
}

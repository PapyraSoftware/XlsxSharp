namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// A redundant prefix <c>=</c> directly in front of an expression, e.g. <c>IF(=7,1,0)</c> - Excel
/// (via Lotus 1-2-3 legacy compatibility) tolerates a leading <c>=</c> anywhere an expression can
/// start, not just once at the very beginning of the whole formula (the top-level one is normally
/// stripped as plain text before the formula ever reaches a parser - see
/// <c>XlsxSharp.Excel.CalcEngine.FormulaParser.GetAst</c> - but that trick doesn't reach one nested
/// inside a function argument like this). Its value is exactly the wrapped expression's value -
/// same "no-op wrapper" shape as <see cref="GroupParselet{TScalar,T,TContext}"/> uses for parens,
/// including going through <see cref="IAstFactory{TScalarValue,TNode,TContext}.Nested"/> so a
/// text-reconstructing consumer (e.g. <c>FormulaConverter</c>) can still recover the "=" in its
/// output rather than silently dropping it.
/// </summary>
internal class RedundantEqualsParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public RedundantEqualsParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token equalToken)
    {
        Node<T> inner = this._parser.ParseExpression(ctx, 0);
        SymbolRange range = new(equalToken.Range.Start, inner.Range.End);
        T value = this._factory.Nested(ctx, range, inner.Value);
        return new Node<T>(value, range, inner.IsPureReference);
    }
}

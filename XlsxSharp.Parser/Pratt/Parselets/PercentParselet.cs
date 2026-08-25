namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Postfix <c>%</c> operator, e.g. <c>50%</c> or <c>-2%</c> (which is <c>Percent(Minus(2))</c> -
/// percent wraps the whole preceding unary chain, it isn't part of it).
/// </summary>
internal class PercentParselet<TScalar, T, TContext> : IParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;

    public PercentParselet(IAstFactory<TScalar, T, TContext> factory)
    {
        this._factory = factory;
    }

    public Node<T> Parse(TContext ctx, Node<T> left, Token op)
    {
        // Not left.Range.ExtendRight(op.Range): whitespace can now separate the operand from the
        // "%" (e.g. "1 %"), so strict adjacency no longer holds.
        SymbolRange range = new(left.Range.Start, op.Range.End);
        T node = this._factory.Unary(ctx, range, UnaryOperation.Percent, left);
        return new Node<T>(node, range);
    }

    public int GetBindingPower()
    {
        return BindingPower.Percent;
    }
}

namespace XlsxSharp.Parser.Pratt.Parselets;

internal class BinaryOpParselet<TScalar, T, TContext> : IParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;
    private readonly BinaryOperation _op;
    private readonly int _bp;

    public BinaryOpParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser, BinaryOperation op, int bp)
    {
        this._factory = factory;
        this._parser = parser;
        this._op = op;
        this._bp = bp;
    }

    public Node<T> Parse(TContext ctx, Node<T> left, Token op)
    {
        Node<T> right = this._parser.ParseExpression(ctx, this._bp);

        // Not left.Range.ExtendRight(op.Range).ExtendRight(right.Range): that asserts strict
        // adjacency between the tokens, which no longer holds now that whitespace can separate
        // them (e.g. "1 + 2").
        SymbolRange nodeRange = new(left.Range.Start, right.Range.End);

        T node = this._factory.BinaryNode(ctx, nodeRange, this._op, left, right);
        return new Node<T>(node, nodeRange);
    }

    public int GetBindingPower()
    {
        return this._bp;
    }
}


namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Prefix <c>+</c>/<c>-</c> operator. Recurses on itself (<c>--1</c> is valid), and its operand
/// stops at the bare atom - see the remarks on <see cref="BindingPower.Percent"/> for why a
/// following <c>%</c> or <c>^</c> must not be swallowed into the operand.
/// </summary>
internal class UnaryOpParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;
    private readonly UnaryOperation _op;

    public UnaryOpParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser, UnaryOperation op)
    {
        this._factory = factory;
        this._parser = parser;
        this._op = op;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        Node<T> operand = this._parser.ParseExpression(ctx, BindingPower.Percent);
        SymbolRange range = new(token.Range.Start, operand.Range.End);
        T node = this._factory.Unary(ctx, range, this._op, operand);
        return new Node<T>(node, range);
    }
}

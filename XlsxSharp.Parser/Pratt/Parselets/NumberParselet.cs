using System.Globalization;

namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Get a number node from a <see cref="TokenType.Number"/> token.
/// </summary>
/// <remarks>
/// <c>double.Parse</c> parses even <c>NaN</c> or <c>∞</c>, but we can never receive such text
/// from the lexer.
/// </remarks>
internal class NumberParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public NumberParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
#if NETSTANDARD2_1
        var text = token.GetText(_parser.Input);
#else
        string text = token.GetText(this._parser.Input).ToString(); // NetFx has a double whammy, it's slow and gets extra memory to GC
#endif
        double number = double.Parse(text, NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture);
        T node = this._factory.NumberNode(ctx, token.Range, number);
        return new Node<T>(node, token.Range);
    }
}

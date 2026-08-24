namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// Get a text node from a <see cref="TokenType.Text"/> token, e.g. <c>"abc"</c> or <c>"a""b"</c>
/// (an escaped quote, unescaping to <c>a"b</c>).
/// </summary>
internal class TextParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public TextParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        string text = Unescape(token.GetText(this._parser.Input));
        T node = this._factory.TextNode(ctx, token.Range, text);
        return new Node<T>(node, token.Range);
    }

    /// <summary>
    /// Strip the surrounding double quotes and collapse escaped <c>""</c> pairs into a single
    /// <c>"</c>, e.g. <c>"a""b"</c> becomes <c>a"b</c>.
    /// </summary>
    private static string Unescape(ReadOnlySpan<char> quotedText)
    {
        ReadOnlySpan<char> inner = quotedText[1..^1];
        if (inner.IndexOf('"') < 0)
        {
            return inner.ToString();
        }

        Span<char> buffer = new char[inner.Length];
        int w = 0;
        int i = 0;
        while (i < inner.Length)
        {
            if (inner[i] == '"')
            {
                i++;
            }

            buffer[w++] = inner[i++];
        }

        return buffer[..w].ToString();
    }
}

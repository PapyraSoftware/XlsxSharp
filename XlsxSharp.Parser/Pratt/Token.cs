namespace XlsxSharp.Parser.Pratt;

internal readonly struct Token
{
    public Token(TokenType type, int start, int end)
    {
        this.Type = type;
        this.Range = new SymbolRange(start, end);
    }

    public TokenType Type { get; }

    public SymbolRange Range { get; }

    public ReadOnlySpan<char> GetText(string input)
    {
        return input.AsSpan(this.Range.Start, this.Range.Length);
    }
}

using System.Globalization;

namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// An array literal (<see cref="TokenType.LeftCurly"/>), e.g. <c>{1,2,3}</c> or <c>{1,2;3,4}</c>
/// (two rows). Elements are always literal scalar constants (number, text, logical or error) -
/// never a reference, a function call, or a nested array - and a leading <c>+</c>/<c>-</c> is only
/// allowed directly on a number, not recursively (<c>{--1}</c> isn't valid). All rows must have
/// the same number of columns.
/// </summary>
internal class ArrayParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public ArrayParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token leftCurly)
    {
        List<List<TScalar>> rows = [[]];

        while (true)
        {
            rows[^1].Add(this.ParseScalarConstant(ctx));

            this._parser.SkipWhitespace();
            Token separator = this._parser.Consume();
            if (separator.Type == TokenType.Comma)
            {
                continue;
            }

            if (separator.Type == TokenType.Semicolon)
            {
                rows.Add([]);
                continue;
            }

            if (separator.Type == TokenType.RightCurly)
            {
                int columns = rows[0].Count;
                foreach (List<TScalar> row in rows)
                {
                    if (row.Count != columns)
                    {
                        throw new ParsingException($"Unable to parse value starting from position {leftCurly.Range.Start}.");
                    }
                }

                List<TScalar> elements = [];
                foreach (List<TScalar> row in rows)
                {
                    elements.AddRange(row);
                }

                SymbolRange range = new(leftCurly.Range.Start, separator.Range.End);
                T value = this._factory.ArrayNode(ctx, range, rows.Count, columns, elements);
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {leftCurly.Range.Start}.");
        }
    }

    private TScalar ParseScalarConstant(TContext ctx)
    {
        this._parser.SkipWhitespace();
        Token token = this._parser.Consume();

        if (token.Type is TokenType.Plus or TokenType.Minus)
        {
            Token numberToken = this._parser.Consume();
            if (numberToken.Type != TokenType.Number)
            {
                throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
            }

            double magnitude = ParseNumber(numberToken, this._parser.Input);
            double signed = token.Type == TokenType.Minus ? -magnitude : magnitude;
            return this._factory.NumberValue(ctx, new SymbolRange(token.Range.Start, numberToken.Range.End), signed);
        }

        switch (token.Type)
        {
            case TokenType.Number:
                return this._factory.NumberValue(ctx, token.Range, ParseNumber(token, this._parser.Input));

            case TokenType.Text:
                return this._factory.TextValue(ctx, token.Range, ParserExtensions.UnescapeText(token.GetText(this._parser.Input)));

            case TokenType.Error:
                return this._factory.ErrorValue(ctx, token.Range, token.GetText(this._parser.Input).ToString().ToUpperInvariant());

            case TokenType.Ident when ParserExtensions.EqualCaseInsensitive(token.GetText(this._parser.Input), "TRUE"):
                return this._factory.LogicalValue(ctx, token.Range, true);

            case TokenType.Ident when ParserExtensions.EqualCaseInsensitive(token.GetText(this._parser.Input), "FALSE"):
                return this._factory.LogicalValue(ctx, token.Range, false);

            default:
                throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }
    }

    private static double ParseNumber(Token token, string input)
    {
        string text = token.GetText(input).ToString();
        return double.Parse(text, NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture);
    }
}

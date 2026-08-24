namespace XlsxSharp.Parser.Pratt.Parselets;

internal class IdentParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public IdentParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        // When we receive an ident, there are following possibilities what it could be (checked
        // in this order):
        // * A1:B2
        // * A1
        // * A:B
        // * $4:6 - rowspan starting with an absolute row
        // * sheet!A1:A2
        // * sheet!A1
        // * sheet!A:B
        // * sheet!$1:2
        // * sheet!name
        // * sheet!1:2
        // * TRUE/FALSE
        // * sheet1:sheet2!A1:A2
        // * sheet1:sheet2!A1
        // * sheet1:sheet2!A:B
        // * sheet1:sheet2!$1:2
        // * name

        // Check for area `A1:B2` or just cell `A1`
        // Check for colspan `A:B`
        // Check for colspan `$1:2` with absolute row start, because this is an "ident" prefix parselet
        if (this._parser.TryReferenceA1(token, out ReferenceArea localArea, out SymbolRange localAreaRange))
        {
            T value = this._factory.Reference(ctx, localAreaRange, localArea);
            return new Node<T>(value, localAreaRange);
        }


        if (this._parser.TryGetUnquotedSheet(token, out ReadOnlySpan<char> sheetNameSpan) && this._parser.LookAhead(1).Type == TokenType.Bang)
        {
            // We are now in `sheet!` Parse local reference.
            string sheetName = sheetNameSpan.ToString(); // String allocation, needed for the IAstFactory
            Token bangToken = this._parser.Consume(TokenType.Bang);
            SymbolRange sheetWithBangRange = token.Range.ExtendRight(bangToken.Range);

            // No need to check for token type, if EoF, nothing will be matched to such token
            Token sheetRefToken = this._parser.Consume();
            
            // Check for area `sheet!A1:B2` or just cell `sheet!A1`
            // Check for colspan `sheet!A:B`
            // Check for rowspan `sheet!1:2` with absolute or relative start row
            if (this._parser.TryReferenceA1(sheetRefToken, out ReferenceArea sheetArea, out SymbolRange sheetAreaRange))
            {
                SymbolRange range = sheetWithBangRange.ExtendRight(sheetAreaRange);
                T value = this._factory.SheetReference(ctx, range, sheetName, sheetArea);
                return new Node<T>(value, range);
            }

            // Check for `sheet!name`
            if (this._parser.TryGetName(sheetRefToken, out ReadOnlySpan<char> name))
            {
                SymbolRange range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
                T value = this._factory.SheetName(ctx, range, sheetName, name.ToString()); // String allocation, needed for the IAstFactory
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        ReadOnlySpan<char> tokenText = token.GetText(this._parser.Input);
        if (EqualCaseInsensitive(tokenText, "TRUE"))
        {
            T value = this._factory.LogicalNode(ctx, token.Range, true);
            return new Node<T>(value, token.Range);
        }

        if (EqualCaseInsensitive(tokenText, "FALSE"))
        {
            T value = this._factory.LogicalNode(ctx, token.Range, false);
            return new Node<T>(value, token.Range);
        }

        // Check for 3D reference for unquoted sheets:
        // * Sheet1:Sheet2!A1:B2
        // * Sheet1:Sheet2!A1
        // * Sheet1:Sheet2!A:B
        // * Sheet1:Sheet2!1:2
        if (this._parser.TryGetUnquotedSheet(token, out ReadOnlySpan<char> startSheet) &&
            this._parser.LookAhead(1).Type == TokenType.Range &&
            this._parser.LookAhead(2) is { Type: TokenType.Ident } maybeEndSheetToken &&
            this._parser.TryGetUnquotedSheet(maybeEndSheetToken, out ReadOnlySpan<char> endSheet) &&
            this._parser.LookAhead(3).Type == TokenType.Bang)
        {
            Token sheetStartToken = token;
            Token rangeToken = this._parser.Consume(TokenType.Range);
            Token sheetEndToken = this._parser.Consume(TokenType.Ident);
            Token bangToken = this._parser.Consume(TokenType.Bang);
            Token refToken = this._parser.Consume();

            if (this._parser.TryReferenceA1(refToken, out ReferenceArea sheetRangeReference, out SymbolRange sheetRangeReferenceRange))
            {
                SymbolRange range = sheetStartToken.Range
                    .ExtendRight(rangeToken.Range)
                    .ExtendRight(sheetEndToken.Range)
                    .ExtendRight(bangToken.Range)
                    .ExtendRight(sheetRangeReferenceRange);
                string startSheetString = startSheet.ToString(); // String allocation for the IAstFactory
                string endSheetString = endSheet.ToString();
                T value = this._factory.Reference3D(ctx, range, startSheetString, endSheetString, sheetRangeReference);
                return new Node<T>(value, range);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        // Check for rowspan `name`
        if (this._parser.TryGetName(token, out ReadOnlySpan<char> workbookName))
        {
            T value = this._factory.Name(ctx, token.Range, workbookName.ToString()); // String allocation, needed for the IAstFactory
            return new Node<T>(value, token.Range);
        }

        throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
    }

    private static bool EqualCaseInsensitive(ReadOnlySpan<char> text, string other)
    {
        if (text.Length != other.Length)
        {
            return false;
        }

        return text.CompareTo(other.AsSpan(), StringComparison.OrdinalIgnoreCase) == 0;
    }
}

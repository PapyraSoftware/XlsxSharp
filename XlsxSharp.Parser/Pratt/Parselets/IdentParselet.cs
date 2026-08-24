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
        // * NAME(...) - local function call, or A1(...) - cell function call
        // * A1:B2
        // * A1
        // * A:B
        // * $4:6 - rowspan starting with an absolute row
        // * sheet!NAME(...) - sheet-scoped function call
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

        // Check for a function call `NAME(...)` or a cell function call `A1(...)`. This has to be
        // checked before anything else: e.g. `A1` alone is a reference, but `A1(...)` never is,
        // and `TRUE` alone is a logical, but `TRUE(...)` is a call to a function named TRUE.
        if (this._parser.LookAhead(1).Type == TokenType.LeftParen)
        {
            return this.ParseFunctionCall(ctx, token, sheet: null);
        }

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

            // Check for a sheet-scoped function call `sheet!NAME(...)`. There is no sheet-scoped
            // cell function form in the grammar (only a bare `A1(...)` can call a LAMBDA stored in
            // a cell), so `sheet!A1(...)` is rejected rather than silently misparsed.
            if (this._parser.LookAhead(1).Type == TokenType.LeftParen)
            {
                if (ParserExtensions.TryGetCellA1(sheetRefToken.GetText(this._parser.Input), out _))
                {
                    throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
                }

                return this.ParseFunctionCall(ctx, sheetRefToken, sheetName);
            }

            // Check for `sheet!#REF!` - a reference to a deleted sheet. Only #REF! is special
            // here (other errors, e.g. `sheet!#N/A`, aren't valid); the whole thing collapses to
            // an error. Unlike ErrorParselet, the oracle doesn't normalize the casing in this
            // particular path, so pass the text through as-is to match it exactly.
            if (sheetRefToken.Type == TokenType.Error && EqualCaseInsensitive(sheetRefToken.GetText(this._parser.Input), "#REF!"))
            {
                SymbolRange range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
                T value = this._factory.ErrorNode(ctx, range, sheetRefToken.GetText(this._parser.Input));
                return new Node<T>(value, range);
            }

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

    /// <summary>
    /// Parse a function call, having already seen its name token (a bare local function/cell
    /// function name, or the name that follows a consumed <c>sheet!</c> prefix) and confirmed the
    /// next token is <see cref="TokenType.LeftParen"/>.
    /// </summary>
    private Node<T> ParseFunctionCall(TContext ctx, Token nameToken, string? sheet)
    {
        ReadOnlySpan<char> name = nameToken.GetText(this._parser.Input);
        bool isCellShaped = ParserExtensions.TryGetCellA1(name, out RowCol cell) && sheet is null;

        this._parser.Consume(TokenType.LeftParen);
        (List<T> args, Token rightParen) = this.ParseArgumentList(ctx);
        SymbolRange range = new(nameToken.Range.Start, rightParen.Range.End);

        T value = sheet is null
            ? isCellShaped
                ? this._factory.CellFunction(ctx, range, cell, args)
                : this._factory.Function(ctx, range, name, args)
            : this._factory.Function(ctx, range, sheet, name, args);
        return new Node<T>(value, range);
    }

    /// <summary>
    /// Parse a function call argument list, having already consumed the opening
    /// <see cref="TokenType.LeftParen"/>. Arguments may be blank (e.g. <c>SUM(1,,2)</c>,
    /// <c>SUM(1,)</c>, <c>SUM(,1)</c>), but an entirely empty list (<c>SUM()</c>) has zero
    /// arguments rather than a single blank one.
    /// </summary>
    private (List<T> Args, Token RightParen) ParseArgumentList(TContext ctx)
    {
        List<T> args = [];
        if (this._parser.LookAhead(1).Type == TokenType.RightParen)
        {
            return (args, this._parser.Consume(TokenType.RightParen));
        }

        while (true)
        {
            Token next = this._parser.LookAhead(1);
            if (next.Type == TokenType.Comma)
            {
                Token comma = this._parser.Consume(TokenType.Comma);
                args.Add(this._factory.BlankNode(ctx, new SymbolRange(comma.Range.Start, comma.Range.Start)));
                continue;
            }

            if (next.Type == TokenType.RightParen)
            {
                Token rightParen = this._parser.Consume(TokenType.RightParen);
                args.Add(this._factory.BlankNode(ctx, new SymbolRange(rightParen.Range.Start, rightParen.Range.Start)));
                return (args, rightParen);
            }

            Node<T> arg = this._parser.ParseExpression(ctx, 0);
            args.Add(arg.Value);

            if (this._parser.LookAhead(1).Type == TokenType.RightParen)
            {
                return (args, this._parser.Consume(TokenType.RightParen));
            }

            this._parser.Consume(TokenType.Comma);
        }
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

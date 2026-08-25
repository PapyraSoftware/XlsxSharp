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
            return new Node<T>(value, localAreaRange, isPureReference: true);
        }


        if (this._parser.TryGetUnquotedSheet(token, out ReadOnlySpan<char> sheetNameSpan) && this._parser.LookAhead(1).Type == TokenType.Bang)
        {
            return this.ParseSheetQualified(ctx, token, sheetNameSpan.ToString());
        }

        ReadOnlySpan<char> tokenText = token.GetText(this._parser.Input);
        if (ParserExtensions.EqualCaseInsensitive(tokenText, "TRUE"))
        {
            T value = this._factory.LogicalNode(ctx, token.Range, true);
            return new Node<T>(value, token.Range);
        }

        if (ParserExtensions.EqualCaseInsensitive(tokenText, "FALSE"))
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
                return new Node<T>(value, range, isPureReference: true);
            }

            throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
        }

        // Check for a table-qualified structure reference `Table1[Column]`. A bare `[Column]`
        // (no table name) is handled by StructureReferenceParselet instead, since it doesn't
        // start with an ident.
        if (this._parser.LookAhead(1).Type == TokenType.SquareIdent && this._parser.TryGetName(token, out ReadOnlySpan<char> tableName))
        {
            Token squareIdentToken = this._parser.Consume(TokenType.SquareIdent);
            TokenParser.ParseIntraTableReference(
                squareIdentToken.GetText(this._parser.Input),
                out StructuredReferenceArea area,
                out string? firstColumn,
                out string? lastColumn
            );
            SymbolRange range = new(token.Range.Start, squareIdentToken.Range.End);
            T structureReferenceValue = this._factory.StructureReference(ctx, range, tableName.ToString(), area, firstColumn, lastColumn ?? firstColumn);
            return new Node<T>(structureReferenceValue, range, isPureReference: true);
        }

        // Check for rowspan `name`
        if (this._parser.TryGetName(token, out ReadOnlySpan<char> workbookName))
        {
            T value = this._factory.Name(ctx, token.Range, workbookName.ToString()); // String allocation, needed for the IAstFactory
            return new Node<T>(value, token.Range, isPureReference: true);
        }

        throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
    }

    /// <summary>
    /// Parse everything after an unquoted sheet name has already been confirmed (see
    /// <see cref="ParserExtensions.TryGetUnquotedSheet{T,TContext}"/>) and is immediately followed
    /// by <c>!</c>. Shared between a bare NAME-shaped sheet prefix (this class's own
    /// <see cref="Parse"/>) and a purely numeric one (<see cref="NumberParselet{TScalar,T,TContext}"/>,
    /// e.g. <c>6!A1</c>) - Excel allows an all-digit sheet name, and the oracle's
    /// SINGLE_SHEET_PREFIX lexer token doesn't care which token type the digits would otherwise
    /// have lexed as on their own.
    /// </summary>
    internal Node<T> ParseSheetQualified(TContext ctx, Token token, string sheetName)
    {
        Token bangToken = this._parser.Consume(TokenType.Bang);
        SymbolRange sheetWithBangRange = token.Range.ExtendRight(bangToken.Range);

        // No need to check for token type, if EoF, nothing will be matched to such token
        Token sheetRefToken = this._parser.Consume();

        // Check for a sheet-scoped function call `sheet!NAME(...)`. There is no sheet-scoped
        // cell function form in the grammar (only a bare `A1(...)` can call a LAMBDA stored in
        // a cell), so `sheet!A1(...)` is rejected rather than silently misparsed.
        if (this._parser.LookAhead(1).Type == TokenType.LeftParen)
        {
            if (this._parser.TryGetCell(sheetRefToken.GetText(this._parser.Input), out _))
            {
                throw new ParsingException($"Unable to parse value starting from position {token.Range.Start}.");
            }

            return this.ParseFunctionCall(ctx, sheetRefToken, sheetName);
        }

        // Check for `sheet!#REF!` - a reference to a deleted sheet. Only #REF! is special
        // here (other errors, e.g. `sheet!#N/A`, aren't valid); the whole thing collapses to
        // an error. Unlike ErrorParselet, the oracle doesn't normalize the casing in this
        // particular path, so pass the text through as-is to match it exactly.
        if (sheetRefToken.Type == TokenType.Error && ParserExtensions.EqualCaseInsensitive(sheetRefToken.GetText(this._parser.Input), "#REF!"))
        {
            SymbolRange range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
            T value = this._factory.ErrorNode(ctx, range, sheetRefToken.GetText(this._parser.Input));
            return new Node<T>(value, range, isPureReference: true);
        }

        // Check for area `sheet!A1:B2` or just cell `sheet!A1`
        // Check for colspan `sheet!A:B`
        // Check for rowspan `sheet!1:2` with absolute or relative start row
        if (this._parser.TryReferenceA1(sheetRefToken, out ReferenceArea sheetArea, out SymbolRange sheetAreaRange))
        {
            SymbolRange range = sheetWithBangRange.ExtendRight(sheetAreaRange);
            T value = this._factory.SheetReference(ctx, range, sheetName, sheetArea);
            return new Node<T>(value, range, isPureReference: true);
        }

        // Check for `sheet!name`
        if (this._parser.TryGetName(sheetRefToken, out ReadOnlySpan<char> name))
        {
            SymbolRange range = sheetWithBangRange.ExtendRight(sheetRefToken.Range);
            T value = this._factory.SheetName(ctx, range, sheetName, name.ToString()); // String allocation, needed for the IAstFactory
            return new Node<T>(value, range, isPureReference: true);
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
        bool isCellShaped = this._parser.TryGetCell(name, out RowCol cell) && sheet is null;

        this._parser.Consume(TokenType.LeftParen);
        (List<T> args, Token rightParen) = this._parser.ParseArgumentList(this._factory, ctx);
        SymbolRange range = new(nameToken.Range.Start, rightParen.Range.End);

        T value = sheet is null
            ? isCellShaped
                ? this._factory.CellFunction(ctx, range, cell, args)
                : this._factory.Function(ctx, range, name, args)
            : this._factory.Function(ctx, range, sheet, name, args);

        // Only an unqualified call to one of the oracle's five "reference functions" (CHOOSE, IF,
        // INDEX, INDIRECT, OFFSET) can be an operand of the range operator (":") - e.g.
        // "INDEX(A1:A5,1):C3" is valid, but "Sheet1!INDEX(...)":C3" and "SUM(...):C3" aren't.
        bool isPureReference = sheet is null && !isCellShaped && ParserExtensions.IsRefFunctionName(name);
        return new Node<T>(value, range, isPureReference);
    }
}

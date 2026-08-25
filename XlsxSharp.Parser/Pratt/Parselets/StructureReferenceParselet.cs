using System.Globalization;

namespace XlsxSharp.Parser.Pratt.Parselets;

/// <summary>
/// A <see cref="TokenType.SquareIdent"/> not immediately preceded by a name is either:
/// <list type="bullet">
///   <item>A bare structure reference with no table name, e.g. <c>[Column]</c> or
///   <c>[#Totals]</c> - only valid for a formula entered directly in the table (e.g. a totals
///   row). A table-qualified structure reference (<c>Table1[Column]</c>) is handled by
///   <see cref="IdentParselet{TScalar,T,TContext}"/> instead, since it starts with a name.</item>
///   <item>A bare external workbook index, e.g. <c>[2]</c> in <c>[2]Sheet!A1</c> or <c>[2]!name</c>
///   - the oracle's lexer never produces a table-less INTRA_TABLE_REFERENCE whose bracket content
///   is just digits, so that shape unambiguously means "workbook index" instead.</item>
/// </list>
/// </summary>
/// <remarks>
/// Delegates the actual bracket-content parsing to <see cref="TokenParser.ParseIntraTableReference"/>,
/// the same routine <see cref="FormulaParser{TScalarValue,TNode,TContext}"/> uses - the pratt
/// lexer's <see cref="TokenType.SquareIdent"/> token covers exactly the same bracket syntax
/// (including up to two levels of nested brackets), so there's no reason to duplicate that logic.
/// </remarks>
internal class StructureReferenceParselet<TScalar, T, TContext> : IPrefixParselet<T, TContext>
{
    private readonly IAstFactory<TScalar, T, TContext> _factory;
    private readonly Parser<T, TContext> _parser;

    public StructureReferenceParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser)
    {
        this._factory = factory;
        this._parser = parser;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        if (TryGetWorkbookIndex(token.GetText(this._parser.Input), out int workbookIndex))
        {
            return this.ParseExternalReference(ctx, token, workbookIndex);
        }

        TokenParser.ParseIntraTableReference(
            token.GetText(this._parser.Input),
            out StructuredReferenceArea area,
            out string? firstColumn,
            out string? lastColumn
        );
        T value = this._factory.StructureReference(ctx, token.Range, area, firstColumn, lastColumn ?? firstColumn);
        return new Node<T>(value, token.Range, isPureReference: true);
    }

    private static bool TryGetWorkbookIndex(ReadOnlySpan<char> squareIdentText, out int workbookIndex)
    {
        ReadOnlySpan<char> inner = squareIdentText[1..^1];
        if (inner.Length == 0)
        {
            workbookIndex = 0;
            return false;
        }

        foreach (char c in inner)
        {
            if (!CompatUtils.IsAsciiDigit(c))
            {
                workbookIndex = 0;
                return false;
            }
        }

        return int.TryParse(inner, NumberStyles.None, CultureInfo.InvariantCulture, out workbookIndex);
    }

    /// <summary>
    /// Parse an external workbook reference, having already seen the bare workbook index prefix
    /// (e.g. <c>[2]</c>). Only forms actually accepted by the oracle are implemented: an external
    /// structure reference (<c>[2]Table1[Column]</c>) has no grammar production and is rejected by
    /// the oracle too, so it isn't handled here either. Whitespace anywhere in this construct
    /// (e.g. between <c>]</c> and <c>!</c>, or between the sheet name and <c>!</c>) is rejected by
    /// the oracle as well, so every lookahead/consume below is intentionally raw (not whitespace
    /// tolerant).
    /// </summary>
    private Node<T> ParseExternalReference(TContext ctx, Token workbookToken, int workbookIndex)
    {
        // [n]!name or [n]!func(...) - no sheet.
        if (this._parser.LookAhead(1).Type == TokenType.Bang)
        {
            this._parser.Consume(TokenType.Bang);
            Token nameToken = this._parser.Consume();

            if (this._parser.LookAhead(1).Type == TokenType.LeftParen)
            {
                return this.ParseExternalFunctionCall(ctx, workbookToken, workbookIndex, sheet: null, nameToken);
            }

            if (this._parser.TryGetName(nameToken, out ReadOnlySpan<char> name) && !this._parser.TryGetCell(name, out _))
            {
                SymbolRange range = new(workbookToken.Range.Start, nameToken.Range.End);
                T value = this._factory.ExternalName(ctx, range, workbookIndex, name.ToString());
                return new Node<T>(value, range, isPureReference: true);
            }

            throw new ParsingException($"Unable to parse value starting from position {workbookToken.Range.Start}.");
        }

        // [n]Sheet!... or [n]Sheet1:Sheet2!... - sheet-qualified forms.
        if (this._parser.LookAhead(1).Type == TokenType.Ident)
        {
            Token sheetToken = this._parser.Consume(TokenType.Ident);

            // [n]Sheet1:Sheet2!A1
            if (
                this._parser.LookAhead(1).Type == TokenType.Range
                && this._parser.LookAhead(2).Type == TokenType.Ident
                && this._parser.LookAhead(3).Type == TokenType.Bang
            )
            {
                this._parser.Consume(TokenType.Range);
                Token sheetEndToken = this._parser.Consume(TokenType.Ident);
                this._parser.Consume(TokenType.Bang);
                Token refToken = this._parser.Consume();

                if (this._parser.TryReferenceA1(refToken, out ReferenceArea area3D, out SymbolRange area3DRange))
                {
                    SymbolRange range = new(workbookToken.Range.Start, area3DRange.End);
                    T value = this._factory.ExternalReference3D(
                        ctx,
                        range,
                        workbookIndex,
                        sheetToken.GetText(this._parser.Input).ToString(),
                        sheetEndToken.GetText(this._parser.Input).ToString(),
                        area3D
                    );
                    return new Node<T>(value, range, isPureReference: true);
                }

                throw new ParsingException($"Unable to parse value starting from position {workbookToken.Range.Start}.");
            }

            if (this._parser.LookAhead(1).Type == TokenType.Bang)
            {
                this._parser.Consume(TokenType.Bang);
                string sheet = sheetToken.GetText(this._parser.Input).ToString();

                Token refOrNameToken = this._parser.Consume();

                if (this._parser.LookAhead(1).Type == TokenType.LeftParen)
                {
                    return this.ParseExternalFunctionCall(ctx, workbookToken, workbookIndex, sheet, refOrNameToken);
                }

                // [n]Sheet!#REF! - a reference to a deleted sheet, same as the local sheet!#REF!
                // case (see IdentParselet).
                if (
                    refOrNameToken.Type == TokenType.Error
                    && ParserExtensions.EqualCaseInsensitive(refOrNameToken.GetText(this._parser.Input), "#REF!")
                )
                {
                    SymbolRange range = new(workbookToken.Range.Start, refOrNameToken.Range.End);
                    T value = this._factory.ErrorNode(ctx, range, refOrNameToken.GetText(this._parser.Input));
                    return new Node<T>(value, range, isPureReference: true);
                }

                if (this._parser.TryReferenceA1(refOrNameToken, out ReferenceArea sheetArea, out SymbolRange sheetAreaRange))
                {
                    SymbolRange range = new(workbookToken.Range.Start, sheetAreaRange.End);
                    T value = this._factory.ExternalSheetReference(ctx, range, workbookIndex, sheet, sheetArea);
                    return new Node<T>(value, range, isPureReference: true);
                }

                if (this._parser.TryGetName(refOrNameToken, out ReadOnlySpan<char> name))
                {
                    SymbolRange range = new(workbookToken.Range.Start, refOrNameToken.Range.End);
                    T value = this._factory.ExternalSheetName(ctx, range, workbookIndex, sheet, name.ToString());
                    return new Node<T>(value, range, isPureReference: true);
                }

                throw new ParsingException($"Unable to parse value starting from position {workbookToken.Range.Start}.");
            }
        }

        throw new ParsingException($"Unable to parse value starting from position {workbookToken.Range.Start}.");
    }

    /// <summary>
    /// Parse an external function call, having already seen its name token and confirmed the next
    /// token is <see cref="TokenType.LeftParen"/>. Unlike a local/sheet-qualified function name, a
    /// cell-shaped external function name (e.g. <c>[2]!A1(1)</c>) is rejected by the oracle even
    /// though it would be unambiguous - there's no external cell function form in the grammar.
    /// </summary>
    private Node<T> ParseExternalFunctionCall(TContext ctx, Token workbookToken, int workbookIndex, string? sheet, Token nameToken)
    {
        ReadOnlySpan<char> name = nameToken.GetText(this._parser.Input);
        if (this._parser.TryGetCell(name, out _))
        {
            throw new ParsingException($"Unable to parse value starting from position {workbookToken.Range.Start}.");
        }

        this._parser.Consume(TokenType.LeftParen);
        (List<T> args, Token rightParen) = this._parser.ParseArgumentList(this._factory, ctx);
        SymbolRange range = new(workbookToken.Range.Start, rightParen.Range.End);

        T value = sheet is null
            ? this._factory.ExternalFunction(ctx, range, workbookIndex, name, args)
            : this._factory.ExternalFunction(ctx, range, workbookIndex, sheet, name, args);
        return new Node<T>(value, range);
    }
}

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
    private readonly IdentParselet<TScalar, T, TContext> _identParselet;

    public NumberParselet(IAstFactory<TScalar, T, TContext> factory, Parser<T, TContext> parser, IdentParselet<TScalar, T, TContext> identParselet)
    {
        this._factory = factory;
        this._parser = parser;
        this._identParselet = identParselet;
    }

    public Node<T> Parse(TContext ctx, Token token)
    {
        // A bare digit-only row span, e.g. "1:2" (as opposed to "$4:6", which starts with "$" and
        // so lexes as an Ident, handled by IdentParselet instead). TryReferenceA1 only matches here
        // when a valid ":<row>" continuation actually follows - a plain number like "1" alone falls
        // through to the normal numeric literal parsing below.
        if (this._parser.TryReferenceA1(token, out ReferenceArea area, out SymbolRange areaRange))
        {
            T reference = this._factory.Reference(ctx, areaRange, area);
            return new Node<T>(reference, areaRange, isPureReference: true);
        }

        // A purely numeric sheet name, e.g. "6!A1" - Excel allows an all-digit sheet name, and the
        // oracle's SINGLE_SHEET_PREFIX lexer token matches it unquoted the same as any other short
        // simple sheet name (its character class doesn't exclude digits). Note this deliberately
        // doesn't go through ParserExtensions.TryGetUnquotedSheet (unlike every other unquoted-
        // sheet check in this parser): that helper also requires !NameUtils.ShouldQuote, which is a
        // writer-side "quote this for clarity" policy that flags every digit-first name - the
        // oracle's own reader-side grammar has no such restriction, so gating on it here would
        // reject formulas the oracle accepts. Delegate to IdentParselet's shared sheet-qualified
        // parsing rather than duplicating it (function call, #REF!, area, name - everything after
        // "sheet!" is identical either way).
        ReadOnlySpan<char> sheetNameSpan = token.GetText(this._parser.Input);
        if (NameUtils.IsSheetNameValid(sheetNameSpan) && this._parser.LookAhead(1).Type == TokenType.Bang)
        {
            return this._identParselet.ParseSheetQualified(ctx, token, sheetNameSpan.ToString());
        }

        ReadOnlySpan<char> text = token.GetText(this._parser.Input);
        double number = double.Parse(text, NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture);
        T node = this._factory.NumberNode(ctx, token.Range, number);
        return new Node<T>(node, token.Range);
    }
}

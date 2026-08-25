using JetBrains.Annotations;
using XlsxSharp.Parser.Pratt;
using XlsxSharp.Parser.Pratt.Parselets;

namespace XlsxSharp.Parser;

/// <summary>
/// A utility class that parses various types of references.
/// </summary>
public static class ReferenceParser
{
    /// <summary>
    /// <para>
    /// Try to parse <paramref name="text"/> as a sheet reference (<c>Sheet!A5</c>) or a local
    /// reference (<c>A1</c>). If the <paramref name="text"/> is a local reference, the output
    /// value of the <paramref name="sheetName"/> is <c>null</c>.
    /// </para>
    /// <para>
    /// Unlike the <see cref="TryParseA1(string,out ReferenceArea)"/> or <see cref="TryParseSheetA1(string, out string, out ReferenceArea)"/>,
    /// this method can parse both sheet reference or local reference.
    /// </para>
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="sheetName">The unescaped name of a sheet for sheet reference, <c>null</c> for local reference.</param>
    /// <param name="area">The parsed reference area.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseA1(string text, out string? sheetName, out ReferenceArea area)
    {
        if (text is null)
        {
            throw new ArgumentNullException();
        }

        sheetName = null;
        if (!TryTokenize(text, out Pratt.Token[] tokens))
        {
            area = default;
            return false;
        }

        if (TryParseBareReference(text, tokens, out area))
        {
            return true;
        }

        if (TrySplitSheetPrefix(text, tokens, out int? workbookIndex, out string sheet, out int start) &&
            workbookIndex is null &&
            TryParseBareReference(text, tokens.AsSpan(start), out area))
        {
            sheetName = sheet;
            return true;
        }

        area = default;
        return false;
    }

    /// <summary>
    /// Parses area reference in A1 form. The possibilities are
    /// <list type="bullet">
    ///   <item>Cell (e.g. <c>F8</c>).</item>
    ///   <item>Area (e.g. <c>B2:$D7</c>).</item>
    ///   <item>Colspan (e.g. <c>$D:$G</c>).</item>
    ///   <item>Rowspan (e.g. <c>14:$15</c>).</item>
    /// </list>
    /// Doesn't allow any whitespaces or extra values inside.
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="area">Parsed area.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseA1(string text, out ReferenceArea area)
    {
        if (text is null)
        {
            throw new ArgumentNullException();
        }

        if (!TryTokenize(text, out Pratt.Token[] tokens))
        {
            area = default;
            return false;
        }

        return TryParseBareReference(text, tokens, out area);
    }

    /// <summary>
    /// Parses area reference in A1 form. The possibilities are
    /// <list type="bullet">
    ///   <item>Cell (e.g. <c>F8</c>).</item>
    ///   <item>Area (e.g. <c>B2:$D7</c>).</item>
    ///   <item>Colspan (e.g. <c>$D:$G</c>).</item>
    ///   <item>Rowspan (e.g. <c>14:$15</c>).</item>
    /// </list>
    /// Doesn't allow any whitespaces or extra values inside.
    /// </summary>
    /// <exception cref="ParsingException">Invalid input.</exception>
    [PublicAPI]
    public static ReferenceArea ParseA1(string text)
    {
        if (!TryParseA1(text, out ReferenceArea area))
        {
            throw new ParsingException($"Unable to parse '{text}'.");
        }

        return area;
    }

    /// <summary>
    /// Try to parse a A1 reference that has a sheet (e.g. <c>'Data values'!A$1:F10</c>).
    /// If <paramref name="text"/> contains only reference without a sheet or anything
    /// else (e.g. <c>A1</c>), return <c>false</c>.
    /// </summary>
    /// <remarks>
    /// The method doesn't accept
    /// <list type="bullet">
    ///   <item>Sheet names, e.g. <c>Sheet!name</c>.</item>
    ///   <item>External sheet references, e.g. <c>[1]Sheet!A1</c>.</item>
    ///   <item>Sheet errors, e.g. <c>Sheet5!$REF!</c>.</item>
    /// </list>
    /// </remarks>
    /// <param name="text">Text to parse.</param>
    /// <param name="sheetName">Name of the sheet, unescaped (e.g. the sheetName will contain <c>Jane's</c> for <c>'Jane''s'!A1</c>).</param>
    /// <param name="area">Parsed reference.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseSheetA1(string text, out string sheetName, out ReferenceArea area)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        if (!TryTokenize(text, out Pratt.Token[] tokens) ||
            !TrySplitSheetPrefix(text, tokens, out int? workbookIndex, out sheetName, out int start) ||
            workbookIndex is not null ||
            !TryParseBareReference(text, tokens.AsSpan(start), out area))
        {
            sheetName = string.Empty;
            area = default;
            return false;
        }

        return true;
    }

    /// <summary>
    /// <para>
    /// Try to parse <paramref name="text"/> as a name (e.g. <c>Name</c>) or a sheet name
    /// (<c>Sheet!Name</c>). If the <paramref name="text"/> is only a name, the output value of the
    /// <paramref name="sheetName"/> is <c>null</c>.
    /// </para>
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="sheetName">The unescaped name of a sheet for sheet name, <c>null</c> for a name.</param>
    /// <param name="name">The parsed name.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseName(string text, out string? sheetName, out string name)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        sheetName = null;
        if (!TryTokenize(text, out Pratt.Token[] tokens))
        {
            name = string.Empty;
            return false;
        }

        if (TryParseBareName(text, tokens, out name))
        {
            return true;
        }

        if (TrySplitSheetPrefix(text, tokens, out int? workbookIndex, out string sheet, out int start) &&
            workbookIndex is null &&
            TryParseBareName(text, tokens.AsSpan(start), out name))
        {
            sheetName = sheet;
            return true;
        }

        name = string.Empty;
        return false;
    }

    /// <summary>
    /// Try to parse a text as a sheet name (e.g. <c>Sheet!Name</c>). Doesn't accept pure name
    /// without sheet (e.g. <c>name</c>).
    /// </summary>
    /// <param name="text">Text to parse.</param>
    /// <param name="sheetName">Parsed sheet name, unescaped.</param>
    /// <param name="name">Parsed defined name.</param>
    /// <returns><c>true</c> if parsing was a success, <c>false</c> otherwise.</returns>
    [PublicAPI]
    public static bool TryParseSheetName(string text, out string sheetName, out string name)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        if (!TryTokenize(text, out Pratt.Token[] tokens) ||
            !TrySplitSheetPrefix(text, tokens, out int? workbookIndex, out sheetName, out int start) ||
            workbookIndex is not null ||
            !TryParseBareName(text, tokens.AsSpan(start), out name))
        {
            sheetName = string.Empty;
            name = string.Empty;
            return false;
        }

        return true;
    }

    /// <summary>
    /// Tokenize <paramref name="text"/> with the pratt lexer, including the trailing
    /// <see cref="TokenType.Eof"/> token. Returns <c>false</c> (rather than throwing) for text the
    /// lexer can't tokenize at all (e.g. an unterminated quoted literal, an unpaired surrogate) -
    /// every caller here treats that the same as "this text doesn't match the shape it's looking
    /// for".
    /// </summary>
    private static bool TryTokenize(string text, out Pratt.Token[] tokens)
    {
        try
        {
            Lexer lexer = new(text);
            List<Pratt.Token> buffer = [];
            Pratt.Token token;
            do
            {
                token = lexer.Consume();
                buffer.Add(token);
            } while (token.Type != TokenType.Eof);

            tokens = [.. buffer];
            return true;
        }
        catch (ParsingException)
        {
            tokens = [];
            return false;
        }
    }

    /// <summary>
    /// Recognize a sheet prefix - <c>Name!</c> or <c>'Name'!</c>, optionally itself prefixed with
    /// an external workbook index (<c>[2]</c>) - at the start of <paramref name="tokens"/>, and
    /// decode it via <see cref="TokenParser.ParseSingleSheetPrefix"/>. That method only ever looks
    /// at a text span (originally the span of the RDS/Rolex lexer's fused SINGLE_SHEET_PREFIX
    /// token), never at token metadata, so reconstructing the exact substring it would have fused -
    /// from the start of the sheet-name token to the end of the following "!" token - reuses its
    /// unescaping and workbook-index detection unchanged, even though the pratt lexer tokenizes the
    /// sheet name and "!" separately.
    /// </summary>
    /// <remarks>
    /// Returns <c>false</c> - a sheet prefix isn't recognized at all - for an <em>unquoted</em>
    /// external prefix such as <c>[2]Sheet!...</c>: "[" lexes as a <see cref="TokenType.SquareIdent"/>
    /// in the pratt lexer, never fusing with the identifier that follows it, so <c>tokens[0]</c>
    /// there is never <see cref="TokenType.Ident"/>/<see cref="TokenType.QIdent"/>. Every caller
    /// wants that text rejected either way (see e.g. <see cref="TryParseSheetA1"/>'s remarks on
    /// external sheet references), so non-recognition and explicit rejection are observably the
    /// same outcome.
    /// </remarks>
    private static bool TrySplitSheetPrefix(string text, Pratt.Token[] tokens, out int? workbookIndex, out string sheetName, out int remainingStart)
    {
        if (tokens.Length < 2 || tokens[0].Type is not (TokenType.Ident or TokenType.QIdent) || tokens[1].Type != TokenType.Bang)
        {
            workbookIndex = null;
            sheetName = string.Empty;
            remainingStart = 0;
            return false;
        }

        ReadOnlySpan<char> sheetPrefixSpan = text.AsSpan(tokens[0].Range.Start, tokens[1].Range.End - tokens[0].Range.Start);
        TokenParser.ParseSingleSheetPrefix(sheetPrefixSpan, out workbookIndex, out sheetName);
        remainingStart = 2;
        return true;
    }

    /// <summary>
    /// Does <paramref name="tokens"/> consist of exactly one reference-shaped token, or exactly two
    /// reference-shaped tokens joined by a bare <see cref="TokenType.Range"/> - and nothing else
    /// (no whitespace, no leftover content)? A single token only counts as a match if it's a full
    /// cell: a bare column or row alone is only reference-shaped paired with a matching corner via
    /// ":" (mirrors <c>ParserExtensions.TryLocalColSpanA1</c>/<c>TryLocalRowSpanA1</c>, which
    /// require the same colon continuation to fire at all).
    /// </summary>
    private static bool TryParseBareReference(string text, ReadOnlySpan<Pratt.Token> tokens, out ReferenceArea area)
    {
        if (tokens.Length == 2 && tokens[0].Type is TokenType.Ident or TokenType.Number && tokens[1].Type == TokenType.Eof)
        {
            if (ParserExtensions.TryGetCellA1(tokens[0].GetText(text), out RowCol cell))
            {
                area = new ReferenceArea(cell);
                return true;
            }

            area = default;
            return false;
        }

        if (tokens.Length == 4 &&
            tokens[0].Type is TokenType.Ident or TokenType.Number &&
            tokens[1].Type == TokenType.Range &&
            tokens[2].Type is TokenType.Ident or TokenType.Number &&
            tokens[3].Type == TokenType.Eof)
        {
            ReadOnlySpan<char> left = tokens[0].GetText(text);
            ReadOnlySpan<char> right = tokens[2].GetText(text);

            if (ParserExtensions.TryGetCellA1(left, out RowCol cell1) && ParserExtensions.TryGetCellA1(right, out RowCol cell2))
            {
                area = new ReferenceArea(cell1, cell2);
                return true;
            }

            if (ParserExtensions.TryGetColA1(left, out RowCol col1) && ParserExtensions.TryGetColA1(right, out RowCol col2))
            {
                area = new ReferenceArea(col1, col2);
                return true;
            }

            if (ParserExtensions.TryGetRowA1(left, out RowCol row1) && ParserExtensions.TryGetRowA1(right, out RowCol row2))
            {
                area = new ReferenceArea(row1, row2);
                return true;
            }
        }

        area = default;
        return false;
    }

    /// <summary>
    /// Does <paramref name="tokens"/> consist of exactly one name-shaped identifier and nothing
    /// else? Excludes TRUE/FALSE (never a valid name, matching <c>ParserExtensions.TryGetName</c>)
    /// and anything full-cell-shaped (e.g. "A1"): the RDS/Rolex lexer this replaces distinguished a
    /// NAME token from an A1_CELL token at the lexer level, so a cell-shaped identifier was never
    /// even a candidate name to begin with - the pratt lexer only has one generic Ident token, so
    /// that priority has to be re-applied by hand here (a bare column/row alone, e.g. "A", is
    /// deliberately *not* excluded: it's only reference-shaped when paired with a ":" partner, same
    /// as <see cref="TryParseBareReference"/>'s single-token case).
    /// </summary>
    private static bool TryParseBareName(string text, ReadOnlySpan<Pratt.Token> tokens, out string name)
    {
        if (tokens.Length == 2 && tokens[0].Type == TokenType.Ident && tokens[1].Type == TokenType.Eof)
        {
            ReadOnlySpan<char> span = tokens[0].GetText(text);
            if (NameUtils.IsNameValid(span) &&
                !ParserExtensions.EqualCaseInsensitive(span, "TRUE") &&
                !ParserExtensions.EqualCaseInsensitive(span, "FALSE") &&
                !ParserExtensions.TryGetCellA1(span, out _))
            {
                name = span.ToString();
                return true;
            }
        }

        name = string.Empty;
        return false;
    }
}

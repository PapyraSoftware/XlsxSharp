using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// R1C1-mode-specific tokenization of the pratt <see cref="Lexer"/>: a single reference (row,
/// column, or cell) must come out as one <see cref="TokenType.Ident"/> token, including the
/// bracketed relative form (<c>R[-1]</c>), which - unlike everything else the lexer treats as part
/// of an identifier - isn't itself made of ident-continuation characters. Mirrors the shapes
/// exercised by <see cref="R1C1ReferenceTokenTests"/> against the RDS/Rolex lexer.
/// </summary>
public class R1C1LexerTests
{
    [Arguments("R")]
    [Arguments("C")]
    [Arguments("R5")]
    [Arguments("C5")]
    [Arguments("R1048576")]
    [Arguments("C16384")]
    [Arguments("R[0]")]
    [Arguments("R[-14]")]
    [Arguments("C[-14]")]
    [Arguments("R[14]")]
    [Arguments("RC")]
    [Arguments("R1C1")]
    [Arguments("R[7]C2")]
    [Arguments("R812C[7]")]
    [Arguments("R[-1]C[-1]")]
    [Arguments("R[-1]C[-2]")]
    [Test]
    public async Task R1C1ReferenceIsSingleIdentToken(string input)
    {
        Lexer lexer = new(input, isR1C1: true);
        Pratt.Token token = lexer.Consume();
        await Assert.That(token.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(token.GetText(input).ToString()).IsEqualTo(input);
        await Assert.That(lexer.Consume().Type).IsEqualTo(TokenType.Eof);
    }

    // Longer NAME always wins over a shorter R1C1-shaped prefix - matches the oracle's own
    // maximal-munch lexer (see FormulaLexer.g4's commented-out R1C1 grammar block).
    [Arguments("Revenue")]
    [Arguments("Costs")]
    [Arguments("Row1")]
    [Arguments("Column5")]
    [Arguments("R1C1style")]
    [Arguments("RC_Name")]
    [Arguments("R1Comment")]
    [Test]
    public async Task NameLongerThanR1C1PrefixWinsWholeToken(string input)
    {
        Lexer lexer = new(input, isR1C1: true);
        Pratt.Token token = lexer.Consume();
        await Assert.That(token.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(token.GetText(input).ToString()).IsEqualTo(input);
    }

    // A table-qualified structure reference must still lex as Ident + SquareIdent in R1C1 mode,
    // not get swallowed as if "[123]" were an R1C1 relative-offset bracket.
    [Test]
    public async Task StructureReferenceBracketIsNotFusedIntoIdent()
    {
        Lexer lexer = new("Cost[Column1]", isR1C1: true);
        Pratt.Token identToken = lexer.Consume();
        await Assert.That(identToken.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(identToken.GetText("Cost[Column1]").ToString()).IsEqualTo("Cost");

        Pratt.Token squareToken = lexer.Consume();
        await Assert.That(squareToken.Type).IsEqualTo(TokenType.SquareIdent);
        await Assert.That(squareToken.GetText("Cost[Column1]").ToString()).IsEqualTo("[Column1]");
    }

    // Same source text lexes differently depending on mode: in A1 mode, "R[-1]C[-1]" is four
    // separate tokens (no R1C1-bracket fusion at all).
    [Test]
    public async Task BracketFusionOnlyAppliesInR1C1Mode()
    {
        const string input = "R[-1]C[-1]";
        Lexer lexer = new(input, isR1C1: false);

        Pratt.Token r = lexer.Consume();
        await Assert.That(r.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(r.GetText(input).ToString()).IsEqualTo("R");

        Pratt.Token bracket1 = lexer.Consume();
        await Assert.That(bracket1.Type).IsEqualTo(TokenType.SquareIdent);
        await Assert.That(bracket1.GetText(input).ToString()).IsEqualTo("[-1]");

        Pratt.Token c = lexer.Consume();
        await Assert.That(c.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(c.GetText(input).ToString()).IsEqualTo("C");

        Pratt.Token bracket2 = lexer.Consume();
        await Assert.That(bracket2.Type).IsEqualTo(TokenType.SquareIdent);
        await Assert.That(bracket2.GetText(input).ToString()).IsEqualTo("[-1]");
    }

    [Test]
    public async Task UnclosedR1C1BracketFallsBackAndStillFailsAsSquareIdent()
    {
        // "R[abc]" isn't a valid relative offset (no digits), so it must NOT be fused - "R" lexes
        // alone, and "[abc]" is left for ordinary SquareIdent scanning to accept or reject.
        Lexer lexer = new("R[abc]", isR1C1: true);
        Pratt.Token r = lexer.Consume();
        await Assert.That(r.Type).IsEqualTo(TokenType.Ident);
        await Assert.That(r.GetText("R[abc]").ToString()).IsEqualTo("R");

        Pratt.Token square = lexer.Consume();
        await Assert.That(square.Type).IsEqualTo(TokenType.SquareIdent);
    }
}

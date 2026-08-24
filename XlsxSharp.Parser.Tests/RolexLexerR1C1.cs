using XlsxSharp.Parser.Rolex;

namespace XlsxSharp.Parser.Tests;

public class RolexLexerR1C1
{
    [Test]
    [Arguments("r1c1")]
    [Arguments("r[0]c[0]")]
    [Arguments("rc")] // Degenerate R[0]C[0]
    public async Task IgnoresCaseInReferences(string formula)
    {
        List<Token> tokens = RolexLexer.GetTokensR1C1(formula);
        await Assert.That(tokens.Count).IsEqualTo(2);
        await Assert.That(tokens[0].SymbolId).IsEqualTo(Token.A1_CELL);
    }

    [Test]
    [Arguments("R1C1")]
    [Arguments("R[1]C[1]")]
    [Arguments("R1C[1]")]
    [Arguments("R[1]C1")]
    [Arguments("R[1048575]C1")]
    [Arguments("R1C[16383]")]
    [Arguments("R[-1048575]C1")]
    [Arguments("R1C[-16383]")]
    public async Task AbsoluteAndRelativeReferencesCanBeCombined(string formula)
    {
        List<Token> tokens = RolexLexer.GetTokensR1C1(formula);
        await Assert.That(tokens.Count).IsEqualTo(2);
        await Assert.That(tokens[0].SymbolId).IsEqualTo(Token.A1_CELL);
    }

    [Test]
    [Arguments(
        "R[1048576]C1",
        Token.A1_SPAN_REFERENCE,
        Token.INTRA_TABLE_REFERENCE,
        Token.A1_SPAN_REFERENCE,
        Token.EofSymbolId
    )]
    [Arguments(
        "R[-1048576]C1",
        Token.A1_SPAN_REFERENCE,
        Token.INTRA_TABLE_REFERENCE,
        Token.A1_SPAN_REFERENCE,
        Token.EofSymbolId
    )]
    [Arguments("R1C[16384]", Token.A1_CELL, Token.INTRA_TABLE_REFERENCE, Token.EofSymbolId)]
    [Arguments("R1C[-16384]", Token.A1_CELL, Token.INTRA_TABLE_REFERENCE, Token.EofSymbolId)]
    public async Task RelativeReferencesCantReachOutsideOfWorksheet(
        string formula,
        params int[] expectedSymbols
    )
    {
        // Because relative references are one off, i.e. at row 1, the R[1] references second row
        // they can't have full range of columns and rows.
        List<Token> tokens = RolexLexer.GetTokensR1C1(formula);
        await Assert.That(tokens.Select(x => x.SymbolId)).IsEquivalentTo(expectedSymbols);
    }

    [Test]
    public async Task AstralCodePointsAreValidStrings()
    {
        // Grinning face code point U+1F600
        List<Token> tokens = RolexLexer.GetTokensR1C1("\"\uD83D\uDE00\"");
        await Assert
            .That(tokens.Select(x => x.SymbolId))
            .IsEquivalentTo(new[] { Token.STRING_CONSTANT, Token.EofSymbolId });
    }
}

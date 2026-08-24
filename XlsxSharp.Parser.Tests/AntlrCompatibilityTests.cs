using XlsxSharp.Parser.Rolex;

namespace XlsxSharp.Parser.Tests;

/// <summary>
/// ANTLR parser is the source of truth. This test class checks that ANTLR output and Rolex/RDP have same output.
/// </summary>
public class AntlrCompatibilityTests
{
    [Test]
    [Arguments("./data/enron/formulas.csv")]
    [Arguments("./data/euses/formulas.csv")]
    [Arguments("./data/contributions/formulas.csv")]
    public async Task ProduceSameTokensForDataSets(string dataSetFile)
    {
        foreach (string formula in DataSets.ReadCsv(dataSetFile))
        {
            IReadOnlyList<Token> antlrTokens = AssertFormula.GetAntlrTokens(formula);
            List<Token> rolexTokens = RolexLexer.GetTokensA1(formula.AsSpan());

            await Assert.That(rolexTokens).IsEquivalentTo(antlrTokens);
        }
    }
}

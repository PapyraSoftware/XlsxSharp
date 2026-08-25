namespace XlsxSharp.Parser.Tests.Rules;

public class FormulaRuleTests
{
    [Test]
    public async Task AdditionalTextAfterExpressionIsError()
    {
        await AssertFormula.CheckParsingErrorContains(
            "A1)",
            "Formula wasn't fully consumed, unexpected token RightParen at position 2."
        );
    }
}

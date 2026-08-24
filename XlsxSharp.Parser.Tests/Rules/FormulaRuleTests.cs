namespace XlsxSharp.Parser.Tests.Rules;

public class FormulaRuleTests
{
    [Test]
    public async Task AdditionalTextAfterExpressionIsError()
    {
        await AssertFormula.CheckParsingErrorContains(
            "A1)",
            "The formula `A1)` wasn't parsed correctly. The expression `A1` was parsed, but the rest `)` wasn't."
        );
    }
}

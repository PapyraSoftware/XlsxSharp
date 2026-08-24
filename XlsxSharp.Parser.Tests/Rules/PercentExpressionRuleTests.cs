namespace XlsxSharp.Parser.Tests.Rules;

public class PercentExpressionRuleTests
{
    [Test]
    public async Task ThereCanBeMultiplePercentOperators()
    {
        UnaryNode expectedNode = new(UnaryOperation.Percent)
        {
            Children = [new UnaryNode(UnaryOperation.Percent) { Children = [new ValueNode(1234)] }],
        };
        await AssertFormula.SingleNodeParsed("1234%%", expectedNode);
    }
}

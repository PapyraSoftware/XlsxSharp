namespace XlsxSharp.Parser.Tests.Rules;

public class PrefixAtomExpressionRuleTests
{
    [Test]
    [Arguments("++1", UnaryOperation.Plus)]
    [Arguments("--1", UnaryOperation.Minus)]
    public async Task MultipleUnaryOperators(string formula, UnaryOperation op)
    {
        UnaryNode expectedNode = new(op)
        {
            Children = [new UnaryNode(op) { Children = [new ValueNode(1.0)] }],
        };

        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }
}

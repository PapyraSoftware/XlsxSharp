namespace XlsxSharp.Parser.Tests.Rules;

public class RefImplicitExpressionRuleTests
{
    [Test]
    public async Task ImplicitIntersectionOperatorHasLowerPriorityThanRange()
    {
        UnaryNode expectedNode = new(
            UnaryOperation.ImplicitIntersection,
            new BinaryNode(
                BinaryOperation.Range,
                new ReferenceNode(new ReferenceArea(new RowCol(1, 1, A1), new RowCol(2, 1, A1))),
                new ReferenceNode(new ReferenceArea(3, 1, A1))
            )
        );
        await AssertFormula.SingleNodeParsed("@A1:A2:A3", expectedNode);
    }
}

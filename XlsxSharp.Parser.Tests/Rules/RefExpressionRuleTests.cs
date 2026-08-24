namespace XlsxSharp.Parser.Tests.Rules;

public class RefExpressionRuleTests
{
    [Test]
    [MethodDataSource(nameof(TestCases))]
    public async Task RefExpressionCanHaveMultipleRefIntersectionExpressions(
        string formula,
        AstNode expectedNode
    )
    {
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    public async Task UnionOperatorHasLowerPriorityThanIntersection()
    {
        BinaryNode expectedNode = new(
            BinaryOperation.Union,
            new BinaryNode(
                BinaryOperation.Intersection,
                new ReferenceNode(new ReferenceArea(1, 1, A1)),
                new ReferenceNode(new ReferenceArea(2, 1, A1))
            ),
            new BinaryNode(
                BinaryOperation.Intersection,
                new ReferenceNode(new ReferenceArea(3, 1, A1)),
                new ReferenceNode(new ReferenceArea(4, 1, A1))
            )
        );
        await AssertFormula.SingleNodeParsed("A1 A2,A3 A4", expectedNode);
    }

    [Test]
    public async Task WhitespacesAtTheEndOfFormulaAreIgnored()
    {
        BinaryNode expectedNode = new(
            BinaryOperation.Union,
            new NameNode("some_name"),
            new ReferenceNode(new ReferenceArea(2, 1, A1))
        );
        await AssertFormula.SingleNodeParsed(" some_name , A2 ", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_expression : ref_intersection_expression
            yield return ["A1", new ReferenceNode(new ReferenceArea(1, 1, A1))];

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return
            [
                "A1,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new ReferenceNode(new ReferenceArea(1, 1, A1)),
                    new ReferenceNode(new ReferenceArea(2, 1, A1))
                ),
            ];

            // ref_expression : ref_intersection_expression COMMA ref_intersection_expression
            yield return
            [
                "A1,#REF!,A2",
                new BinaryNode(
                    BinaryOperation.Union,
                    new BinaryNode(
                        BinaryOperation.Union,
                        new ReferenceNode(new ReferenceArea(1, 1, A1)),
                        new ValueNode("Error", "#REF!")
                    ),
                    new ReferenceNode(new ReferenceArea(2, 1, A1))
                ),
            ];
        }
    }
}

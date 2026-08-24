namespace XlsxSharp.Parser.Tests.Rules;

public class RefIntersectionExpressionRuleTests
{
    [Test]
    [MethodDataSource(nameof(TestCases))]
    public async Task HasOneOrMoreElementsSeparatedBySpace(string formula, AstNode expectedNode)
    {
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    public async Task IntersectionOperatorHasLowerPriorityThanImplicitIntersection()
    {
        UnaryNode expectedNode = new(
            UnaryOperation.ImplicitIntersection,
            new BinaryNode(
                BinaryOperation.Intersection,
                new ReferenceNode(new ReferenceArea(new RowCol(1, 1, A1), new RowCol(10, 1, A1))),
                new ReferenceNode(new ReferenceArea(5, 1, A1))
            )
        );
        await AssertFormula.SingleNodeParsed("@A1:A10 A5", expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // ref_intersection_expression : ref_range_expression
            yield return ["A1", new ReferenceNode(new ReferenceArea(1, 1, A1))];

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return
            [
                "A1 A2",
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new ReferenceNode(new ReferenceArea(1, 1, A1)),
                    new ReferenceNode(new ReferenceArea(2, 1, A1))
                ),
            ];

            // ref_intersection_expression : ref_range_expression SPACE ref_range_expression
            yield return
            [
                " A1   A2   A3  ",
                new BinaryNode(
                    BinaryOperation.Intersection,
                    new BinaryNode(
                        BinaryOperation.Intersection,
                        new ReferenceNode(new ReferenceArea(1, 1, A1)),
                        new ReferenceNode(new ReferenceArea(2, 1, A1))
                    ),
                    new ReferenceNode(new ReferenceArea(3, 1, A1))
                ),
            ];
        }
    }
}

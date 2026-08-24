namespace XlsxSharp.Parser.Tests.Rules;

public class ArgumentListRuleTests
{
    [Test]
    [MethodDataSource(nameof(TestCases))]
    public async Task ArgumentListCanHaveZeroOrMoreArguments(string formula, AstNode expectedNode)
    {
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    [Arguments("FUN(1,(A1,two),3)")]
    [Arguments("FUN(1,((A1,two)),3)")]
    [Arguments("FUN( 1 , ( ( A1 , two ) ) , 3 ) ")]
    public async Task ArgumentListInterpretersCommaAsArgumentSeparatorButNestedExpressionInterpretsCommaAsRangeUnionOperator(
        string formula
    )
    {
        FunctionNode expectedNode = new("FUN")
        {
            Children =
            [
                new ValueNode(1),
                new BinaryNode(
                    BinaryOperation.Union,
                    new ReferenceNode(new ReferenceArea(Relative, 1, Relative, 1, A1)),
                    new NameNode("two")
                ),
                new ValueNode(3),
            ],
        };
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // argument_list : CLOSE_BRACE
            yield return ["FUN()", new FunctionNode("FUN") { Children = [] }];

            // argument_list : arg_expression CLOSE_BRACE
            yield return
            [
                "FUN(TRUE)",
                new FunctionNode("FUN") { Children = [new ValueNode("Logical", true)] },
            ];
            yield return
            [
                "FUN(1.5)",
                new FunctionNode("FUN") { Children = [new ValueNode("Number", 1.5)] },
            ];

            // argument_list : arg_expression CLOSE_BRACE
            // arg_expression is not a value directly, but another node
            yield return
            [
                "FUN(100%)",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new UnaryNode(UnaryOperation.Percent)
                        {
                            Children = [new ValueNode("Number", 100.0)],
                        },
                    ],
                },
            ];

            // argument_list : COMMA CLOSE_BRACE
            yield return
            [
                "FUN(,)",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Blank", string.Empty),
                    ],
                },
            ];

            // argument_list : COMMA COMMA CLOSE_BRACE
            yield return
            [
                "FUN(  ,  , )",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Blank", string.Empty),
                    ],
                },
            ];

            // argument_list : arg_expression COMMA COMMA arg_expression CLOSE_BRACE
            yield return
            [
                "FUN(  TRUE ,  , 1.0 )",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new ValueNode("Logical", true),
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Number", 1.0),
                    ],
                },
            ];

            // argument_list : COMMA arg_expression COMMA CLOSE_BRACE
            yield return
            [
                "FUN(   , TRUE , )",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new ValueNode("Blank", string.Empty),
                        new ValueNode("Logical", true),
                        new ValueNode("Blank", string.Empty),
                    ],
                },
            ];

            // argument_list : arg_expression COMMA arg_expression COMMA arg_expression CLOSE_BRACE
            yield return
            [
                "FUN( FALSE  , TRUE , 1 )",
                new FunctionNode("FUN")
                {
                    Children =
                    [
                        new ValueNode("Logical", false),
                        new ValueNode("Logical", true),
                        new ValueNode("Number", 1.0),
                    ],
                },
            ];
        }
    }
}

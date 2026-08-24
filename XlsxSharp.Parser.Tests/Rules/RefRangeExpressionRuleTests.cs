namespace XlsxSharp.Parser.Tests.Rules;

public class RefRangeExpressionRuleTests
{
    [Test]
    [MethodDataSource(nameof(TestData))]
    public async Task HasOneOrMoreElementsSeparatedBySpace(string formula, AstNode expectedNode)
    {
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    public async Task CellAreNotMistakenlyRecognizedAs3dReference()
    {
        string formula = "A1:Sheet1!B2";
        BinaryNode expectedNode = new(BinaryOperation.Range)
        {
            Children =
            [
                new ReferenceNode(new ReferenceArea(1, 1, A1)),
                new SheetReferenceNode("Sheet1", new ReferenceArea(2, 2, A1)),
            ],
        };
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    public async Task ColumnsCanBeSheetNamesFor3dReference()
    {
        // JAN and DEC are columns in A1 notation
        string formula = "JAN:DEC!B2";
        Reference3DNode expectedNode = new("JAN", "DEC", new ReferenceArea(2, 2, A1));
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    [Test]
    public async Task SpillHasHigherPriorityThanRange()
    {
        BinaryNode expectedNode = new(BinaryOperation.Range)
        {
            Children =
            [
                new ReferenceNode(new ReferenceArea(5, 1, A1)),
                new UnaryNode(UnaryOperation.SpillRange) { Children = [new NameNode("Name")] },
            ],
        };

        await AssertFormula.SingleNodeParsed("A5:Name#", expectedNode);
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            // ref_range_expression : ref_atom_expression
            yield return ["A1", new ReferenceNode(new ReferenceArea(1, 1, A1))];

            // ref_range_expression : ref_atom_expression COLON ref_atom_expression
            yield return
            [
                "first:second",
                new BinaryNode(
                    BinaryOperation.Range,
                    new NameNode("first"),
                    new NameNode("second")
                ),
            ];

            // ref_range_expression : ref_atom_expression COLON ref_atom_expression COLON ref_atom_expression
            yield return
            [
                // Parser eats A1:B2 as a single token
                "#REF!:B1:last",
                new BinaryNode(
                    BinaryOperation.Range,
                    new BinaryNode(
                        BinaryOperation.Range,
                        new ValueNode("Error", "#REF!"),
                        new ReferenceNode(new ReferenceArea(1, 2, A1))
                    ),
                    new NameNode("last")
                ),
            ];

            // ref_range_expression : A1_CELL COLON NAME COLON A1_CELL
            yield return
            [
                "A5:B6C7:D8", // Make sure parser doesn't mistake first part of formula for area A5:B6
                new BinaryNode(
                    BinaryOperation.Range,
                    new BinaryNode(
                        BinaryOperation.Range,
                        new ReferenceNode(new ReferenceArea(5, 1, A1)),
                        new NameNode("B6C7")
                    ),
                    new ReferenceNode(new ReferenceArea(8, 4, A1))
                ),
            ];
        }
    }
}

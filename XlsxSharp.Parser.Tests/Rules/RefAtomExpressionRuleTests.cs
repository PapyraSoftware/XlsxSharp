namespace XlsxSharp.Parser.Tests.Rules;

public class RefAtomExpressionRuleTests
{
    [Test]
    public async Task RefError()
    {
        await VerifyNode("#REF!", new ValueNode("Error", "#REF!"));
    }

    [Test]
    [Arguments("#REF!A1")]
    [Arguments("#REF!$B$7")]
    [Arguments("#REF!A1:B5")]
    [Arguments("#REF!$D7:G$15")]
    [Arguments("#REF!7:15")]
    [Arguments("#REF!$7:$15")]
    [Arguments("#REF!B:AG")]
    [Arguments("#REF!$ABC:$ADD")]
    [Arguments("Sheet!#REF!")]
    [Arguments("#REF!#REF!")]
    public async Task RefErrorWithReference(string refError)
    {
        await VerifyNode(refError, new ValueNode("Error", "#REF!"));
    }

    [Test]
    public async Task NestedRefExpression()
    {
        await VerifyNode("((#REF!))", new ValueNode("Error", "#REF!"));
    }

    [Test]
    public async Task CellReference()
    {
        await VerifyNode("A1", new ReferenceNode(new ReferenceArea(1, 1, A1)));
    }

    [Test]
    public async Task RefFunctionCall()
    {
        await VerifyNode(
            "IF(TRUE,B5)",
            new FunctionNode("IF")
            {
                Children = [new ValueNode(true), new ReferenceNode(new ReferenceArea(5, 2, A1))],
            }
        );
    }

    [Test]
    public async Task NameReference()
    {
        await VerifyNode("some_name", new NameNode("some_name"));
    }

    [Test]
    public async Task NameReferenceStarts3DReference()
    {
        await VerifyNode(
            "Sheet1:Sheet3!A1",
            new Reference3DNode("Sheet1", "Sheet3", new ReferenceArea(1, 1, A1))
        );
    }

    [Test]
    public async Task BangReference()
    {
        await VerifyNode("!$A7:$D$9", new BangReferenceNode(ReferenceParser.ParseA1("$A7:$D$9")));
    }

    [Test]
    public async Task BangReferenceError()
    {
        await VerifyNode("!#REF!", new ValueNode("Error", "#REF!"));
    }

    [Test]
    public async Task SheetReferenceCanHaveWhitespaceAfterExclamationMark()
    {
        await VerifyNode(
            "Sheet2!  A1",
            new SheetReferenceNode("Sheet2", new ReferenceArea(1, 1, A1))
        );
    }

    [Test]
    public async Task StructureReference()
    {
        await VerifyNode(
            "Table[Column]",
            new StructureReferenceNode("Table", StructuredReferenceArea.None, "Column", "Column")
        );
    }

    [Test]
    public async Task NestedCantBeNonRef()
    {
        await AssertFormula.CheckParsingErrorContains(
            "(1),#REF!",
            "Formula wasn't fully consumed, unexpected token Comma at position 3."
        );
    }

    [Test]
    public async Task NonRefFunctionCantBeRefAtom()
    {
        await AssertFormula.CheckParsingErrorContains(
            "FUNC(),#REF!",
            "Formula wasn't fully consumed, unexpected token Comma at position 6."
        );
    }

    private static async Task VerifyNode(string formula, AstNode node)
    {
        // Force the formula into a ref_atom_expression through ref_range_expression.
        // ref_range_expression: ref_atom_expression (COLON ref_atom_expression)*
        string adjustedFormula = $"#REF!:{formula}";
        BinaryNode expected = new(BinaryOperation.Range, new ValueNode("Error", "#REF!"), node);
        await AssertFormula.SingleNodeParsed(adjustedFormula, expected);
    }
}

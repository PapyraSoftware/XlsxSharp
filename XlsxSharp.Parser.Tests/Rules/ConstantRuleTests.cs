namespace XlsxSharp.Parser.Tests.Rules;

public class ConstantRuleTests
{
    [Test]
    [Arguments("#DIV/0!")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#VALUE!")]
    [Arguments("#GETTING_DATA")]
    public async Task NonRefErrors(string error)
    {
        await AssertFormula.SingleNodeParsed(error, new ValueNode("Error", error));
        await AssertFormula.SingleNodeParsed(
            error.ToLowerInvariant(),
            new ValueNode("Error", error)
        );
    }

    [Test]
    [Arguments("TRUE", true)]
    [Arguments("FALSE", false)]
    public async Task LogicalConstant(string formula, bool value)
    {
        await AssertFormula.SingleNodeParsed(formula, new ValueNode("Logical", value));
        await AssertFormula.SingleNodeParsed(
            formula.ToLowerInvariant(),
            new ValueNode("Logical", value)
        );
    }

    [Test]
    [Arguments("1.5e2", 150.0)]
    [Arguments("25.0e-2", 0.25)]
    [Arguments("1", 1.0)]
    [Arguments("5.4", 5.4)]
    public async Task NumericalConstant(string formula, double value)
    {
        await AssertFormula.SingleNodeParsed(formula, new ValueNode("Number", value));
        await AssertFormula.SingleNodeParsed(
            formula.ToUpperInvariant(),
            new ValueNode("Number", value)
        );
    }

    [Test]
    [Arguments("\"Hello\"", "Hello")]
    [Arguments("\"Tom \"\"Ben\"\"\"", "Tom \"Ben\"")]
    [Arguments("\"\"", "")]
    public async Task StringConstant(string formula, string text)
    {
        await AssertFormula.SingleNodeParsed(formula, new ValueNode("Text", text));
    }

    [Test]
    public async Task SingleElementArray()
    {
        await AssertFormula.SingleNodeParsed(
            "{1}",
            new ArrayNode(1, 1, new[] { new ScalarValue(1) })
        );
    }

    [Test]
    public async Task ArrayCanContainNumberLogicalTextOrError()
    {
        await AssertFormula.SingleNodeParsed(
            "{ 1.5 , true , \"Test\" , #n/a }",
            new ArrayNode(
                1,
                4,
                new[]
                {
                    new ScalarValue(1.5),
                    new ScalarValue(true),
                    new ScalarValue("Test"),
                    new ScalarValue("Error", "#N/A"),
                }
            )
        );
    }

    [Test]
    public async Task NumberInArrayCanHavePlusPrefix()
    {
        await AssertFormula.SingleNodeParsed(
            "{+3}",
            new ArrayNode(1, 1, new[] { new ScalarValue(3) })
        );
    }

    [Test]
    [Arguments("#REF!")]
    [Arguments("#DIV/0!")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#VALUE!")]
    [Arguments("#GETTING_DATA")]
    public async Task ArrayCanContainErrors(string error)
    {
        await AssertFormula.SingleNodeParsed(
            $"{{{error}}}",
            new ArrayNode(1, 1, new[] { new ScalarValue("Error", error) })
        );
    }

    [Test]
    public async Task ArrayCantContainBlanks()
    {
        await AssertFormula.CheckParsingErrorContains("{1,,}", " Unexpected token COMMA.");
    }

    [Test]
    public async Task EmptyArrayIsUnparsable()
    {
        await AssertFormula.CheckParsingErrorContains("{}", "Unexpected token CLOSE_CURLY.");
    }

    [Test]
    public async Task RowsOfArrayMustHaveSameSize()
    {
        await AssertFormula.CheckParsingErrorContains(
            "{1,2;3,4;5;6,7}",
            "Rows of an array don't have same size."
        );
    }
}

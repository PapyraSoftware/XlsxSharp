namespace XlsxSharp.Parser.Tests.Lexers;

public class ScalarValueTests
{
    [Test]
    public async Task CanParseEmptyString()
    {
        await AssertText("\"\"", string.Empty);
    }

    [Test]
    public async Task CanParseNonEscapedString()
    {
        await AssertText("\"someone's\ntext\"", "someone's\ntext");
    }

    [Test]
    [Arguments("\"\"\"\"", "\"")]
    [Arguments("\"\"\"\"\"\"", "\"\"")]
    [Arguments("\"Eastern \"\"Bonn's\"\" Tavern\"", "Eastern \"Bonn's\" Tavern")]
    public async Task CanParseEscapedString(string unescaped, string escaped)
    {
        await AssertText(unescaped, escaped);
    }

    [Test]
    [Arguments("TRUE", true)]
    [Arguments("FALSE", false)]
    public async Task CanParseLogical(string formula, bool value)
    {
        await AssertValue(formula, "Logical", value);
    }

    [Test]
    [Arguments("#REF!", "#REF!")]
    [Arguments("#N/A", "#N/A")]
    public async Task CanParseError(string formula, string value)
    {
        await AssertValue(formula, "Error", value);
    }

    [Test]
    [Arguments("1", 1)]
    [Arguments("1.5", 1.5)]
    [Arguments(".5", .5)]
    [Arguments(".5E2", 50)]
    // [Arguments(".5e2", 50)] TODO: Lower e
    [Arguments(".5E+2", 50)]
    [Arguments("50E-2", 0.5)]
    public async Task CanParseNumber(string formula, double value)
    {
        await AssertValue(formula, "Number", value);
    }

    private static async Task AssertText<T>(string formula, T expected)
    {
        await AssertValue(formula, "Text", expected);
    }

    private static async Task AssertValue<T>(string formula, string expectedType, T expected)
    {
        ValueNode node = (ValueNode)ParseText(formula, new F());
        await Assert.That(node.Type).IsEqualTo(expectedType);
        await Assert.That(node.Value).IsEqualTo(expected);
    }

    private static AstNode ParseText(string formula, IAstFactory<ScalarValue, AstNode, Ctx> factory)
    {
        return FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), factory);
    }
}

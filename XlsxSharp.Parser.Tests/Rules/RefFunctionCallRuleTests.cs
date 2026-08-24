namespace XlsxSharp.Parser.Tests.Rules;

public class RefFunctionCallRuleTests
{
    [Test]
    public async Task RefFunctionsAreRecognized()
    {
        FunctionNode expected = new("IF") { Children = [new ValueNode("Logical", true)] };
        await AssertFormula.SingleNodeParsed("IF(TRUE)", expected);
    }

    [Test]
    public async Task RefFunctionsCanHaveWhitespacesAroundBraces()
    {
        FunctionNode expected = new("CHOOSE")
        {
            Children = [new ValueNode("Logical", true), new ValueNode("Number", 5.0)],
        };
        await AssertFormula.SingleNodeParsed("CHOOSE(  TRUE, 5  )", expected);
    }
}

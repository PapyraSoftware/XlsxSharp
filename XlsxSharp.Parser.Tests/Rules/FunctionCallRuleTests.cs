namespace XlsxSharp.Parser.Tests.Rules;

public class FunctionCallRuleTests
{
    [Test]
    public async Task PredefinedFunctionsAreRecognized()
    {
        FunctionNode expectedNode = new("SIN") { Children = [new ValueNode("Number", 5.0)] };
        await AssertFormula.SingleNodeParsed("SIN(5)", expectedNode);
    }

    [Test]
    public async Task FunctionCanHaveWhitespacesAroundBraces()
    {
        FunctionNode expectedNode = new("SIN") { Children = [new ValueNode("Number", 5.0)] };
        await AssertFormula.SingleNodeParsed("SIN(  5  )", expectedNode);
    }

    [Test]
    public async Task FunctionCanBeFromAnotherSheet()
    {
        FunctionNode expectedNode = new("Sheet", "Func")
        {
            Children = [new ValueNode("Number", 5.0)],
        };
        await AssertFormula.SingleNodeParsed("Sheet!Func(5)", expectedNode);
    }

    [Test]
    public async Task FunctionCanBeFromAnotherWorkbook()
    {
        ExternalFunctionNode expectedNode = new(2, null, "Func")
        {
            Children = [new ValueNode("Number", 5.0)],
        };
        await AssertFormula.SingleNodeParsed("[2]!Func(5)", expectedNode);
    }

    [Test]
    public async Task FunctionCanBeCellFunction()
    {
        CellFunctionNode expectedNode = new(new RowCol(true, 3, false, 2, A1))
        {
            Children = [new ValueNode("Number", 5.0)],
        };
        await AssertFormula.SingleNodeParsed("B$3(5)", expectedNode);
    }
}

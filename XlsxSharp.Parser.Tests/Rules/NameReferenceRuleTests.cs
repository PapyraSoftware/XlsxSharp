namespace XlsxSharp.Parser.Tests.Rules;

public class NameReferenceRuleTests
{
    [Test]
    public async Task NameIsRecognized()
    {
        NameNode expectedNode = new("SomeName");
        await AssertFormula.SingleNodeParsed("SomeName", expectedNode);
    }

    [Test]
    public async Task SheetNameIsRecognized()
    {
        SheetNameNode expectedNode = new("Sheet", "SomeName");
        await AssertFormula.SingleNodeParsed("Sheet!SomeName", expectedNode);
    }

    [Test]
    public async Task ExternalNameIsRecognized()
    {
        ExternalNameNode expectedNode = new(2, "SomeName");
        await AssertFormula.SingleNodeParsed("[2]!SomeName", expectedNode);
    }

    [Test]
    public async Task ExternalSheetNameIsRecognized()
    {
        ExternalSheetNameNode expectedNode = new(14, "Sheet", "SomeName");
        await AssertFormula.SingleNodeParsed("[14]Sheet!SomeName", expectedNode);
    }
}

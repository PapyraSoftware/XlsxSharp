namespace XlsxSharp.Parser.Tests;

public class FunctionTests
{
    [Test]
    [Arguments("TRUE(TRUE)")]
    public async Task AmbiguousBuiltInFunctionNameIsRecognizedAsFunction(string formula)
    {
        await AssertFormula.CstParsed(formula);
    }
}

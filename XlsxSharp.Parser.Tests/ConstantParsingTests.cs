namespace XlsxSharp.Parser.Tests;

public class ConstantParsingTests
{
    [Test]
    [Arguments("1")]
    [Arguments("10.5")]
    [Arguments("10.5E5")]
    [Arguments(".1E-4")]
    public async Task NumberIsParsed(string formula)
    {
        await AssertFormula.CstParsed(formula);
    }

    [Test]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    public async Task ErrorIsParsed(string formula)
    {
        await AssertFormula.CstParsed(formula);
    }
}

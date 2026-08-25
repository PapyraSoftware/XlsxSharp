using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests;

internal static class AssertFormula
{
    /// <summary>
    /// Assert that a formula is parsed into a single childless node.
    /// </summary>
    public static async Task SingleNodeParsed<TNode>(string formula, TNode expectedNode)
        where TNode : AstNode
    {
        TNode node = (TNode)ParserFactory.Create(new F()).ParseFormula(formula, new Ctx());
        await Assert.That(node).IsEqualTo(expectedNode);
    }

    public static async Task CheckParsingErrorContains(string formula, string errorSubstring)
    {
        Exception ex = Assert.Throws<Exception>(() =>
            ParserFactory.Create(new F()).ParseFormula(formula, new Ctx())
        );
        await Assert
            .That(ex.Message.Contains(errorSubstring))
            .IsTrue()
            .Because($"Error message '{ex.Message}' doesn't contain '{errorSubstring}'.");
    }
}

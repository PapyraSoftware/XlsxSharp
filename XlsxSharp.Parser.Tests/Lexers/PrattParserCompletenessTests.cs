using System.Diagnostics;
using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// The pratt parser must consume the whole formula. A leftover, unclaimed token (a stray closing
/// parenthesis, or an operator/construct that isn't implemented yet, such as a function call)
/// must be reported as a parsing error instead of being silently dropped.
/// </summary>
public class PrattParserCompletenessTests
{
    [Test]
    [Arguments("1+2)")]
    [Arguments("(1+2))")]
    [Arguments("SUM(1,2)")] // Function calls aren't implemented yet, "(1,2)" must not be dropped silently.
    [Arguments("1 1")]
    public async Task TrailingTokensAreRejected(string formula)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsExactlyAsync<ParsingException>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    [Test]
    [Arguments("1+2", "(1+2)")]
    [Arguments("1+2 ", "(1+2)")]
    [Arguments("1+2\t", "(1+2)")]
    public async Task TrailingWhitespaceIsIgnored(string formula, string expectedNormalizedForm)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode root = parser.ParseFormula(formula, new Ctx());
        await Assert.That(GetNormalizedForm(root)).IsEqualTo(expectedNormalizedForm);
    }

    private static string GetNormalizedForm(AstNode node)
    {
        return node switch
        {
            ValueNode value => value.GetDisplayString(A1),
            BinaryNode binaryOp => "("
                + GetNormalizedForm(binaryOp.Children[0])
                + binaryOp.GetDisplayString(A1)
                + GetNormalizedForm(binaryOp.Children[1])
                + ")",
            _ => throw new UnreachableException(),
        };
    }
}

using System.Diagnostics;
using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

public class PrattParserPrecedenceTests
{
    [Test]
    [Arguments("1+2+3+4", "(((1+2)+3)+4)")]
    [Arguments("1-2-3-4", "(((1-2)-3)-4)")]
    [Arguments("1-2+3-4+5", "((((1-2)+3)-4)+5)")]
    [Arguments("1*2*3*4", "(((1*2)*3)*4)")]
    [Arguments("1/2/3/4", "(((1/2)/3)/4)")]
    [Arguments("1*2/3*4/5", "((((1*2)/3)*4)/5)")]
    [Arguments("2^3^4^5", "(((2^3)^4)^5)")] // Even exponential is left-associative in Excel, contrary to standard convention
    public async Task OperationsWithSamePrecedenceAreLeftAssociative(
        string formula,
        string normalizedForm
    )
    {
        await AssertSameFormulas(formula, normalizedForm);
    }

    [Test]
    [Arguments("1+(2+3+4)+((5+6)+7)", "((1+((2+3)+4))+((5+6)+7))")]
    [Arguments("1-(2-3-4)-((5-6)-7)", "((1-((2-3)-4))-((5-6)-7))")]
    [Arguments("1-(2+3-4)+((5-6)+7)", "((1-((2+3)-4))+((5-6)+7))")]
    [Arguments("1*(2*3*4)*((5*6)*7)", "((1*((2*3)*4))*((5*6)*7))")]
    [Arguments("1/(2/3/4)/((5/6)/7)", "((1/((2/3)/4))/((5/6)/7))")]
    [Arguments("1/(2*3/4)*((5/6)*7)", "((1/((2*3)/4))*((5/6)*7))")]
    [Arguments("2^(3^4)^5", "((2^(3^4))^5)")]
    public async Task GroupsOverridePrecedence(string formula, string normalizedForm)
    {
        await AssertSameFormulas(formula, normalizedForm);
    }

    [Test]
    [Arguments("1+2*3+4/5*6^7-8", "(((1+(2*3))+((4/5)*(6^7)))-8)")]
    [Arguments("1+2-3*4+5/6^7-8*9", "((((1+2)-(3*4))+(5/(6^7)))-(8*9))")]
    public async Task OperationsAreGroupedByPrecedence(string formula, string normalizedForm)
    {
        await AssertSameFormulas(formula, normalizedForm);
    }

    private static async Task AssertSameFormulas(string formula, string normalizedForm)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode root = parser.ParseFormula(formula, new Ctx());

        await Assert.That(GetNormalizedForm(root)).IsEqualTo(normalizedForm);
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

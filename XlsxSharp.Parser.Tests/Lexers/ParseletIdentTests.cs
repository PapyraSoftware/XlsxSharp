using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

public class ParseletIdentTests
{
    [Test]
    // Local area
    [Arguments("A1:B1", typeof(ReferenceNode))]
    [Arguments("$A$1:$B$1", typeof(ReferenceNode))]
    // Local cell
    [Arguments("A1", typeof(ReferenceNode))]
    [Arguments("A$1", typeof(ReferenceNode))]
    [Arguments("$A1", typeof(ReferenceNode))]
    [Arguments("$A$1", typeof(ReferenceNode))]
    [Arguments("XFD1048576", typeof(ReferenceNode))]
    [Arguments("XFD$1048576", typeof(ReferenceNode))]
    [Arguments("$XFD1048576", typeof(ReferenceNode))]
    [Arguments("$XFD$1048576", typeof(ReferenceNode))]
    // Local colspan
    [Arguments("A:B", typeof(ReferenceNode))]
    [Arguments("$GE:$XFD", typeof(ReferenceNode))]
    // Local rowspan starting with absolute
    [Arguments("$1:8", typeof(ReferenceNode))]
    [Arguments("$72:$85", typeof(ReferenceNode))]
    // sheet!A1:A2
    [Arguments("Sheet!A1:B2", typeof(SheetReferenceNode))]
    [Arguments("Sheet!$Z$84:$BG$99", typeof(SheetReferenceNode))]
    // sheet!A1
    [Arguments("Sheet!A1", typeof(SheetReferenceNode))]
    [Arguments("Sheet!$Z$84", typeof(SheetReferenceNode))]
    // sheet!$1:2
    [Arguments("Sheet!$4:81", typeof(SheetReferenceNode))]
    [Arguments("Sheet!$1:$5", typeof(SheetReferenceNode))]
    // sheet!name
    [Arguments("Sheet!name", typeof(SheetNameNode))]
    [Arguments("Sheet!_name", typeof(SheetNameNode))]
    // sheet!1:2
    [Arguments("Sheet!1:2", typeof(SheetReferenceNode))]
    [Arguments("Sheet!1:$2", typeof(SheetReferenceNode))]
    // name
    [Arguments("_name", typeof(NameNode))]
    [Arguments("name", typeof(NameNode))]
    // Not a valid reference (leading zero, out of range row/column), but a validly-shaped name -
    // same as the oracle, which doesn't know whether a name is actually defined at parse time.
    [Arguments("A01", typeof(NameNode))]
    [Arguments("A0", typeof(NameNode))]
    [Arguments("A1048577", typeof(NameNode))]
    [Arguments("XFE1", typeof(NameNode))]
    // sheet1:sheet2!A1:B2
    [Arguments("sheet1:sheet2!A1:B2", typeof(Reference3DNode))]
    [Arguments("sheet1:sheet2!$A$1:$B$2", typeof(Reference3DNode))]
    // sheet1:sheet2!A1
    [Arguments("sheet1:sheet2!A1", typeof(Reference3DNode))]
    [Arguments("sheet1:sheet2!$A$1", typeof(Reference3DNode))]
    // sheet1:sheet2!A:B
    [Arguments("sheet1:sheet2!A:C", typeof(Reference3DNode))]
    [Arguments("sheet1:sheet2!$A:$C", typeof(Reference3DNode))]
    // sheet1:sheet2!1:2
    [Arguments("sheet1:sheet2!1:2", typeof(Reference3DNode))]
    [Arguments("sheet1:sheet2!$1:$2", typeof(Reference3DNode))]
    public async Task CanParseReferencesStartingAtIdent(string formula, Type expectedNodeType)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode root = parser.ParseFormula(formula, new Ctx());

        await Assert.That(root.GetType()).IsEqualTo(expectedNodeType);
        await Assert.That(root.GetDisplayString(A1)).IsEqualTo(formula);
    }

    [Test]
    [Arguments("TRUE", true)]
    [Arguments("true", true)]
    [Arguments("FALSE", false)]
    [Arguments("false", false)]
    public async Task CanParseLogical(string formula, bool expectedValue)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode root = parser.ParseFormula(formula, new Ctx());

        await Assert.That(root).IsEqualTo(new ValueNode(expectedValue));
    }

    [Test]
    [Arguments("sheet!$")]
    [Arguments("sheet!")]
    [Arguments("$")]
    [Arguments("sheet1:sheet2!")]
    [Arguments("sheet1:sheet2!A")]
    [Arguments("sheet1:sheet2!name")] // There is no such thing as 3D name
    public void InvalidReferencesStartingWithIdentThrowParsingException(string formula)
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        Assert.ThrowsExactly<ParsingException>(() => parser.ParseFormula(formula, new Ctx()));
    }
}

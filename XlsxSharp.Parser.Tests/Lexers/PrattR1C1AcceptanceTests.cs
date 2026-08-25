using XlsxSharp.Parser.Ast;
using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// R1C1-mode acceptance tests for the pratt parser, mirroring <see cref="PrattParserAcceptanceTests"/>
/// but exercising <see cref="Parser{T,TContext}.ParseFormula"/> with <c>isR1C1: true</c> against
/// <see cref="FormulaParser{TScalarValue,TNode,TContext}.CellFormulaR1C1"/> as the oracle.
/// </summary>
public class PrattR1C1AcceptanceTests
{
    [Arguments("R")]
    [Arguments("C")]
    [Arguments("RC")]
    [Arguments("R5")]
    [Arguments("C5")]
    [Arguments("R1048576")]
    [Arguments("C16384")]
    [Arguments("R[0]")]
    [Arguments("R[-14]")]
    [Arguments("C[-14]")]
    [Arguments("R1C1")]
    [Arguments("R[7]C2")]
    [Arguments("R812C[7]")]
    [Arguments("R[-1]C[-1]")]
    [Arguments("R1C2:R3C4")]
    [Arguments("RC:RC")]
    [Arguments("R[-1]C[-2]:R[-3]C[-4]")]
    [Arguments("R5:R10")]
    [Arguments("C5:C10")]
    [Test]
    public async Task ReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Arguments("R1C1+R2C2")]
    [Arguments("SUM(R1C1:R5C5)")]
    [Arguments("SUM(RC,R1C1)")]
    [Arguments("-R1C1")]
    [Arguments("R1C1%")]
    [Arguments("R1C1&R2C2")]
    [Arguments("IF(R1C1,R2C2,R3C3)")]
    [Test]
    public async Task ReferencesInExpressionsMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Arguments("Sheet1!R1C1")]
    [Arguments("Sheet1!R1C1:R2C2")]
    [Arguments("Sheet1!RC")]
    [Arguments("'New York'!R1C1")]
    // An unquoted sheet name that happens to be R1C1-shaped - row-only, column-only, or a full
    // cell - is still a sheet name, not a reference: NameUtils' quoting policy is A1-shape-aware
    // only, so it never flags these for quoting the way an A1-cell-shaped sheet name would be.
    [Arguments("R6!R1C1")]
    [Arguments("C5!R1C1")]
    [Arguments("R1C1!R2C2")]
    [Arguments("R1C1!RC")]
    [Test]
    public async Task SheetQualifiedReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    // Cell function calls: a full R1C1 cell reference immediately followed by "(" is a call to a
    // LAMBDA stored in that cell, mirroring A1's "A1(...)". Deliberately NOT checked against the
    // oracle here (unlike every other case in this file): TokenParser.ExtractCellFunction always
    // decodes a CELL_FUNCTION_LIST token's name via ReadA1Cell (plain A1 column-letters-then-row-
    // digits), even when the surrounding formula is R1C1 - e.g. for "RC(1,2)" it reads "R"+"C" as
    // a two-letter A1 column ("RC" = 471) with no row digits left, producing garbage (RowValue -8).
    // That's a latent bug in the oracle (R1C1 cell functions are already an obscure corner of an
    // obscure feature, so it's essentially never been exercised), not a behavior worth replicating.
    // Pratt instead decodes the name with the same R1C1 corner-decode used everywhere else, so
    // "R1C1" here produces the same RowCol it would as a plain reference - internally consistent,
    // which the oracle isn't.
    [Test]
    public async Task CellFunctionCallDecodesNameAsR1C1Corner()
    {
        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());

        CellFunctionNode node = (CellFunctionNode)
            parser.ParseFormula("R1C1(1)", new Ctx(), isR1C1: true);
        await Assert
            .That(node.RowCol)
            .IsEqualTo(
                new RowCol(
                    ReferenceAxisType.Absolute,
                    1,
                    ReferenceAxisType.Absolute,
                    1,
                    ReferenceStyle.R1C1
                )
            );

        CellFunctionNode node2 = (CellFunctionNode)
            parser.ParseFormula("R[7]C2(1)", new Ctx(), isR1C1: true);
        await Assert
            .That(node2.RowCol)
            .IsEqualTo(
                new RowCol(
                    ReferenceAxisType.Relative,
                    7,
                    ReferenceAxisType.Absolute,
                    2,
                    ReferenceStyle.R1C1
                )
            );
    }

    // Names that happen to start with R/C but aren't R1C1-shaped remain plain names/functions,
    // matching the oracle's own maximal-munch lexer.
    [Arguments("Revenue")]
    [Arguments("Costs+1")]
    [Arguments("Row1(1)")]
    [Arguments("R1C1style")]
    [Test]
    public async Task NonReferenceNamesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    // A structured/table reference must keep working unchanged in R1C1 mode - "[" after a name
    // isn't mistaken for an R1C1 relative-offset bracket.
    [Arguments("Table1[Column1]")]
    [Arguments("[Column1]")]
    [Test]
    public async Task StructureReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    private static async Task AssertMatchesOracle(string formula)
    {
        AstNode oracleNode = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaR1C1(
            formula,
            new Ctx(),
            new F()
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode prattNode = parser.ParseFormula(formula, new Ctx(), isR1C1: true);

        await Assert.That(prattNode).IsEqualTo(oracleNode);
    }
}

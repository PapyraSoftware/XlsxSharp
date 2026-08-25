using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// Targeted acceptance tests for pratt parser features (function calls, unary/percent operators,
/// comparisons, concatenation) that are broad enough to be awkward to express as a normalized-form
/// string (see <see cref="PrattParserPrecedenceTests"/>). Each case is checked against the same
/// oracle used by <see cref="PrattDataSetTests"/>: the recursive-descent
/// <see cref="FormulaParser{TScalarValue,TNode,TContext}"/>, sharing the same <see cref="F"/>
/// AST factory, so a plain structural equality check is enough.
/// </summary>
public class PrattParserAcceptanceTests
{
    [Test]
    [Arguments("SUM(A1)")]
    [Arguments("SUM(A1,B2)")] // See InsignificantWhitespaceMatchesOracle for the spaced form.
    [Arguments("SUM()")]
    [Arguments("SUM(1,,2)")]
    [Arguments("SUM(,1)")]
    [Arguments("SUM(1,)")]
    [Arguments("IF(TRUE,1,2)")]
    [Arguments("NOT(TRUE)")]
    [Arguments("_xlfn.IFS(1,2)")]
    [Arguments("R1C1(1)")] // "R1C1" isn't a valid A1 cell (trailing letters), so it's a local function.
    [Arguments("R(1)")] // "R" alone has no row digits, so it's a local function too.
    [Arguments("TRUE(1)")] // Followed by "(", TRUE/FALSE are function names, not the logical literal.
    [Arguments("FALSE(1,2)")]
    [Arguments("SUM(-1)")]
    [Arguments("SUM(1)+SUM(2)")]
    [Arguments("A1(1,2)")] // A valid single-cell A1 reference followed by "(" is a cell function.
    [Arguments("Sheet1!SUM(1)")]
    [Arguments("name(1)")]
    [Arguments("_Foo(1)")]
    public async Task FunctionCallsMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("-1")]
    [Arguments("+1")]
    [Arguments("--1")]
    [Arguments("-A1")]
    [Arguments("-A1:B2")]
    [Arguments("1--1")]
    [Arguments("1-+1")]
    [Arguments("1++1")]
    [Arguments("-2^2")] // The famous Excel quirk: unary binds tighter than ^, so this is (-2)^2.
    [Arguments("2^-2")]
    [Arguments("-2^-2")]
    [Arguments("--2^2")]
    [Arguments("-2%")] // Percent wraps the whole unary chain: Percent(Minus(2)), not Minus(Percent(2)).
    [Arguments("-A1%")]
    [Arguments("2%%")]
    [Arguments("2^50%")] // Percent binds tighter than ^: Pow(2, Percent(50)).
    [Arguments("50%^2")]
    [Arguments("-2^2%")]
    [Arguments("-SUM(1)")]
    [Arguments("-SUM(1)%")]
    [Arguments("NOT(-1)")]
    [Arguments("-(1+2)")]
    [Arguments("-(1+2)%")]
    [Arguments("-(1+2)^2")]
    public async Task UnaryAndPercentMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("1+2&3+4")] // Concat is looser than +, tighter than comparisons.
    [Arguments("1&2&3")]
    [Arguments("1=2&3")]
    [Arguments("1<2=3<4")] // Comparisons chain left-associatively, all at the same precedence.
    [Arguments("1<2<3")]
    [Arguments("A1&B1=C1")]
    [Arguments("1+2=3+4")]
    [Arguments("1<>2")]
    [Arguments("1<=2")]
    [Arguments("1>=2")]
    [Arguments("1>2")]
    public async Task ComparisonAndConcatMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("1 + 2")]
    [Arguments("1 +2")]
    [Arguments("1+ 2")]
    [Arguments("1  +  2")] // Multiple spaces collapse to a single Whitespace token either way.
    [Arguments(" 1+2")] // Leading formula whitespace.
    [Arguments("1+2 ")] // Trailing formula whitespace.
    [Arguments("(1 + 2)")]
    [Arguments("( 1 + 2 )")] // Whitespace right after "(" and right before ")".
    [Arguments("SUM(A1, B2)")] // Space after the comma - by far the most common real-world case.
    [Arguments("SUM(A1 , B2)")] // Space before the comma too.
    [Arguments("SUM( A1,B2 )")]
    [Arguments("- 1")] // Space between a prefix operator and its operand.
    [Arguments("-  1")]
    [Arguments("1 %")] // Space between an operand and a postfix operator.
    [Arguments("1%  ^2")]
    [Arguments("SUM(A1) ")]
    [Arguments(" SUM(A1)")]
    public async Task InsignificantWhitespaceMatchesOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("SUM (A1)")] // Whitespace right before "(" still isn't allowed: this is not a call.
    public async Task StillRejectedWhitespaceUsesAreRejectedByBoth(string formula)
    {
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(
                FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F())
            )
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    [Test]
    [Arguments("A1 B1")] // The reference intersection operator isn't implemented yet.
    [Arguments("A1:A10 B1:B10")]
    public async Task ReferenceIntersectionIsNotImplementedYet(string formula)
    {
        // Unlike StillRejectedWhitespaceUsesAreRejectedByBoth, the oracle *does* accept these (as
        // a BinaryNode with BinaryOperation.Intersection) - only the pratt parser doesn't.
        AstNode oracleNode = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(
            formula,
            new Ctx(),
            new F()
        );
        await Assert.That(oracleNode).IsTypeOf<BinaryNode>();
        await Assert
            .That(((BinaryNode)oracleNode).Operation)
            .IsEqualTo(BinaryOperation.Intersection);

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    [Test]
    [Arguments("\"abc\"")]
    [Arguments("\"\"")] // Empty string.
    [Arguments("\"a\"\"b\"")] // Escaped quote: unescapes to a"b.
    [Arguments("\"a\"&\"b\"")]
    [Arguments("SUM(\"1\",\"2\")")]
    [Arguments("\"a\"=\"b\"")]
    public async Task TextLiteralsMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("#REF!")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#DIV/0!")]
    [Arguments("#div/0!")] // Normalized to upper case, regardless of the casing used.
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#GETTING_DATA")]
    [Arguments("SUM(#REF!,1)")]
    [Arguments("#VALUE!+1")]
    [Arguments("Deals!#REF!")] // A reference to a deleted sheet - collapses to a normalized #REF!.
    [Arguments("Deals!#ref!")]
    [Arguments("Deals!#REF!*2")]
    public async Task ErrorLiteralsMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("Deals!#N/A")] // Only #REF! is special-cased after a sheet prefix.
    [Arguments("Deals!#DIV/0!")]
    public async Task SheetPrefixedNonRefErrorsAreRejectedByBoth(string formula)
    {
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(
                FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F())
            )
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    [Test]
    [Arguments("'New York'!A1")]
    [Arguments("'New York'!A1:B2")]
    [Arguments("'New York'!A:B")]
    [Arguments("'New York'!1:2")]
    [Arguments("'Jane''s'!A1")] // '' inside a quoted sheet name is an escaped single quote.
    [Arguments("'Jane''s'!name")]
    [Arguments("'Sheet 1:Sheet 2'!A1")] // The colon here is inside the quotes: a quoted 3D reference.
    [Arguments("'January 1st:December 31st'!A1")]
    [Arguments("1+'Johnny''s'!Z26")]
    public async Task QuotedSheetNamesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("DeptSales[SaleAmt]")] // Table-qualified, single column.
    [Arguments("DeptSales[SaleAmt]*DeptSales[ComPct]")]
    [Arguments("DeptSales[#All]")] // Table-qualified, item specifier keyword.
    [Arguments("DeptSales[#Data]")]
    [Arguments("DeptSales[#Headers]")]
    [Arguments("DeptSales[#Totals]")]
    [Arguments("DeptSales[#This Row]")]
    [Arguments("DeptSales[[#All],[SaleAmt]]")] // Keyword + a single column.
    [Arguments("DeptSales[[#All],[SaleAmt]:[ComPct]]")] // Keyword + a column range.
    [Arguments("DeptSales[[#Headers],[#Data],[ComPct]]")] // Two keywords + a column.
    [Arguments("DeptSales[[SalesPers]:[Region]]")] // Column range, no keyword.
    [Arguments("DeptSales[Total Amount]")] // Column name containing a space.
    [Arguments("[SaleAmt]*[ComPct]")] // No table name (valid only inside the table itself).
    [Arguments("SUBTOTAL(109,[Jan])")]
    [Arguments("SUM(DeptSales[SaleAmt])")]
    [Arguments("VLOOKUP(GroupVertices[[#This Row],[Vertex]],Vertices[],2,FALSE)")] // "[]" is the whole table.
    public async Task StructureReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("SUM (A1)")] // No space is allowed between a function name and "(".
    [Arguments("A1 (1,2)")]
    [Arguments("Sheet1!A1(1,2)")] // No sheet-scoped cell function form exists in the grammar.
    [Arguments("'text'")] // A quoted ident is only ever a sheet name/sheet range prefix, always
    // followed by "!" - a bare one isn't valid anywhere else.
    public async Task RejectedByOracleAreAlsoRejectedByPratt(string formula)
    {
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(
                FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F())
            )
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    [Test]
    [Arguments("'[2]D and D'!A1")] // The workbook index is inside the quotes here.
    public async Task QuotedExternalWorkbookReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("[2]!name")] // No sheet.
    [Arguments("[2]!SUM(1)")]
    [Arguments("[2]Sheet!A1")]
    [Arguments("[2]Sheet!A1:B2")]
    [Arguments("[2]Sheet!A:B")]
    [Arguments("[2]Sheet!1:2")]
    [Arguments("[2]Sheet!name")]
    [Arguments("[2]Sheet!SUM(1)")]
    [Arguments("[2]Sheet1:Sheet2!A1")] // External 3D reference.
    [Arguments("[123]!name")] // Multi-digit workbook index.
    [Arguments("[0]!name")]
    [Arguments("1+[2]Sheet!A1+2")]
    [Arguments("[2]Sheet!#REF!")] // Same #REF! collapsing as the local sheet!#REF! case.
    [Arguments("[2]A1!name")] // "A1" here is a (cell-shaped) sheet name, not a reference/name.
    public async Task UnquotedExternalWorkbookReferencesMatchOracle(string formula)
    {
        await AssertMatchesOracle(formula);
    }

    [Test]
    [Arguments("[2]!A1")] // A cell-shaped name isn't a valid NAME token, external or not.
    [Arguments("[2]!A1(1)")] // Nor a valid external function name - no external cell function form.
    [Arguments("[2]Sheet!A1(1)")]
    [Arguments("[2]TRUE")] // "[2]" alone still needs "!" or a sheet name to follow.
    [Arguments("[2]!TRUE")] // TRUE/FALSE still can't be a bare name, external or not.
    [Arguments("[2]Sheet!TRUE")]
    [Arguments("[2]Table1[Column]")] // No external structure reference form in the grammar.
    public async Task ExternalWorkbookReferenceEdgeCasesAreRejectedByBoth(string formula)
    {
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(
                FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(formula, new Ctx(), new F())
            )
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        await Assert.ThrowsAsync<Exception>(() =>
            Task.FromResult(parser.ParseFormula(formula, new Ctx()))
        );
    }

    private static async Task AssertMatchesOracle(string formula)
    {
        AstNode oracleNode = FormulaParser<ScalarValue, AstNode, Ctx>.CellFormulaA1(
            formula,
            new Ctx(),
            new F()
        );

        Parser<AstNode, Ctx> parser = ParserFactory.Create(new F());
        AstNode prattNode = parser.ParseFormula(formula, new Ctx());

        await Assert.That(prattNode).IsEqualTo(oracleNode);
    }
}

using XlsxSharp.Parser.Pratt;

namespace XlsxSharp.Parser.Tests;

public class RefModVisitorTests
{
    #region ModifySheet

    [Test]
    [Arguments("Old!B7:$D$10", "Old", "New", "New!B7:$D$10")]
    [Arguments("Old!B7:$D$10", "Old", "New sheet", "'New sheet'!B7:$D$10")]
    [Arguments(
        "'Old Mike''s sheet'!B7:$D$10",
        "Old Mike's sheet",
        "New Mike's sheet",
        "'New Mike''s sheet'!B7:$D$10"
    )]
    public async Task ModifySheetCanRenameSheetName(
        string formula,
        string oldSheetName,
        string newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    [Arguments("Old!#REF!", "Old", null, "#REF!#REF!")]
    [Arguments("Old!#REF!", "Old", "New", "New!#REF!")]
    public async Task ErrorNodeCanModifySheet(
        string formula,
        string oldSheetName,
        string? newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    [Arguments("Old!B$5", "Old", null, "#REF!")]
    [Arguments("Old!B:D", "Old", "Shiny", "Shiny!B:D")]
    public async Task SheetReferenceCanModifySheet(
        string formula,
        string oldSheetName,
        string? newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    [Arguments("Old!F(5)", "Old", null, "#REF!")]
    [Arguments("Old!F(7)", "Old", "Shiny", "Shiny!F(7)")]
    public async Task SheetFunctionCanModifySheet(
        string formula,
        string oldSheetName,
        string? newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    [Arguments("Sheet1:Sheet5!A1", "Sheet1", null, "#REF!")]
    [Arguments("Sheet1:Sheet5!A1", "Sheet1", "New sheet", "'New sheet:Sheet5'!A1")]
    [Arguments("Sheet1:Sheet5!A1", "Sheet5", "Sheet9", "Sheet1:Sheet9!A1")]
    public async Task Reference3DCanModifySheet(
        string formula,
        string oldSheetName,
        string? newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    [Arguments("Sheet!Name", "Sheet", null, "#REF!")]
    [Arguments("Sheet!Name", "Sheet", "New Sheet", "'New Sheet'!Name")]
    public async Task SheetNameCanModifySheet(
        string formula,
        string oldSheetName,
        string? newSheetName,
        string modifiedFormula
    )
    {
        FormulaVisitor factory = new() { SheetMap = { { oldSheetName, newSheetName } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    #endregion

    [Test]
    [Arguments("5 + !$B1", "$B1", "$7:$9", "5 + !$7:$9")]
    [Arguments("5 + !$B1", "$B1", null, "5 + !#REF!")]
    public async Task BangReferencesIsModified(
        string formula,
        string reference,
        string? replacement,
        string modifiedFormula
    )
    {
        ShiftReferenceVisitor factory = new() { ReferenceMap = { { reference, replacement } } };
        await AssertChangesA1(formula, factory, modifiedFormula);
    }

    [Test]
    public async Task Log10IsNotInterpretedAsCellFunction()
    {
        ShiftReferenceVisitor factory = new() { ReferenceMap = { { "LOG10", "A1" } } };
        await AssertChangesA1("LOG10(LOG10)", factory, "LOG10(A1)");
    }

    private static async Task AssertChangesA1(
        string formula,
        RefModVisitor visitor,
        string expected
    )
    {
        ModContext ctx = new(formula, "Sheet", 1, 1, isA1: true);
        TransformedSymbol modifiedFormula = ParserFactory
            .Create(visitor)
            .ParseFormula(formula, ctx);
        await Assert.That(modifiedFormula.ToString(string.Empty.AsSpan())).IsEqualTo(expected);
    }

    private class FormulaVisitor : RefModVisitor
    {
        public Dictionary<string, string?> SheetMap { get; } = new();

        protected override string? ModifySheet(ModContext ctx, string sheetName)
        {
            return this.SheetMap.GetValueOrDefault(sheetName, sheetName);
        }
    }

    private class ShiftReferenceVisitor : RefModVisitor
    {
        public Dictionary<string, string?> ReferenceMap { get; } = new();

        internal override ReferenceArea? ModifyRef(ModContext ctx, ReferenceArea reference)
        {
            if (
                this.ReferenceMap.TryGetValue(
                    reference.GetDisplayStringA1(),
                    out string? replacement
                )
            )
            {
                return replacement is not null ? ReferenceParser.ParseA1(replacement) : null;
            }

            return reference;
        }

        internal override RowCol? ModifyCellFunction(ModContext ctx, RowCol cell)
        {
            if (this.ReferenceMap.TryGetValue(cell.GetDisplayStringA1(), out string? replacement))
            {
                return replacement is not null ? ReferenceParser.ParseA1(replacement).First : null;
            }

            return cell;
        }
    }
}

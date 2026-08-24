namespace XlsxSharp.Parser.Tests;

/// <summary>
/// Most of tests are taken care of by <see cref="FormulaConverterToR1C1Tests"/>.
/// </summary>
public class FormulaConverterToA1Tests
{
    [Test]
    [Arguments("RC", 4, 1, "A4")] // References
    [Arguments("R7C", 4, 2, "B$7")]
    [Arguments("R[2]C", 4, 2, "B6")]
    [Arguments("RC5", 4, 2, "$E4")]
    [Arguments("RC[3]", 4, 2, "E4")]
    [Arguments("R[2]C[1]:R[3]C[4]", 4, 3, "D6:G7")]
    [Arguments("C[-2]:C[1]", 4, 3, "A:D")]
    [Arguments("R[-2]:R[6]", 4, 3, "2:10")]
    [Arguments("Sheet4!R[2]C", 4, 2, "Sheet4!B6")] // Sheet reference
    [Arguments("R7C3(TRUE)", 4, 2, "$E$11(TRUE)", Skip = "Parser bug")] // Cell function
    public async Task ExternalSheetReference(string r1c1, int row, int col, string a1)
    {
        await Assert.That(FormulaConverter.ToA1(r1c1, row, col)).IsEqualTo(a1);
    }

    [Test]
    [Arguments("R[-4]C", 4, 1, "#REF!")]
    [Arguments("R[1048575]C", 2, 1, "#REF!")]
    [Arguments("RC[-4]", 1, 4, "#REF!")]
    [Arguments("RC[5]", 1, 16380, "#REF!")]
    public async Task OutOfBoundsReferences(string r1c1, int row, int col, string a1)
    {
        await Assert.That(FormulaConverter.ToA1(r1c1, row, col)).IsEqualTo(a1);
    }

    [Test]
    [Arguments("C2:C4", 1, 1, "$B:$D")]
    [Arguments("C[-2]:C4", 1, 4, "B:$D")]
    [Arguments("C2:C[3]", 1, 4, "$B:G")]
    [Arguments("C[2]:C[3]", 1, 4, "F:G")]
    public async Task ColumnsReference(string r1c1, int row, int col, string a1)
    {
        await Assert.That(FormulaConverter.ToA1(r1c1, row, col)).IsEqualTo(a1);
    }

    [Test]
    [Arguments("R2:R4", 1, 1, "$2:$4")]
    [Arguments("R[-2]:R4", 4, 1, "2:$4")]
    [Arguments("R2:R[3]", 4, 1, "$2:7")]
    [Arguments("R[2]:R[3]", 4, 1, "6:7")]
    public async Task RowsReference(string r1c1, int row, int col, string a1)
    {
        await Assert.That(FormulaConverter.ToA1(r1c1, row, col)).IsEqualTo(a1);
    }
}

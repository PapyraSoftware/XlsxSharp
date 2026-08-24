namespace XlsxSharp.Parser.Tests;

public class RowColTests
{
    [Test]
    public async Task DefaultStructIsA1()
    {
        RowCol a = default;

        await Assert.That(a.ColumnType).IsEqualTo(ReferenceAxisType.Relative);
        await Assert.That(a.ColumnValue).IsEqualTo(1);
        await Assert.That(a.RowType).IsEqualTo(ReferenceAxisType.Relative);
        await Assert.That(a.RowValue).IsEqualTo(1);
        await Assert.That(a.Style).IsEqualTo(A1);
    }

    [Test]
    [Arguments("RC", 1, 1, "A1")]
    [Arguments("RC[-5]", 1, 4, "XFC1")]
    [Arguments("RC[-4]", 1, 4, "XFD1")]
    [Arguments("RC[-3]", 1, 4, "A1")]
    [Arguments("RC[2]", 1, 16382, "XFD1")]
    [Arguments("RC[3]", 1, 16382, "A1")]
    [Arguments("RC[4]", 1, 16382, "B1")]
    [Arguments("R[0]C", 1, 1, "A1")]
    [Arguments("R[-3]C", 4, 1, "A1")]
    [Arguments("R[-4]C", 4, 1, "A1048576")]
    [Arguments("R[-5]C", 4, 1, "A1048575")]
    [Arguments("R[1]C", 1048575, 1, "A1048576")]
    [Arguments("R[2]C", 1048575, 1, "A1")]
    public async Task ToA1LoopsForOutOfBoundsReference(string r1c1, int row, int col, string a1)
    {
        // In GUI, Excel loops over, if user enters out-of-bounds reference to a formula.
        RowCol refR1C1 = TokenParser.ParseReference(r1c1, false).First;
        RowCol refA1 = TokenParser.ParseReference(a1, true).First;
        await Assert.That(refR1C1.ToA1(row, col)).IsEqualTo(refA1);
    }
}

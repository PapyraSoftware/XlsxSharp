namespace XlsxSharp.Parser.Tests;

public class ReferenceTests
{
    [Test]
    [MethodDataSource(nameof(DisplayStringA1))]
    public async Task DisplayStringA1DisplaysReferenceInA1Style(
        RowCol rowCol,
        string expectedString
    )
    {
        await Assert.That(rowCol.GetDisplayStringA1()).IsEqualTo(expectedString);
    }

    [Test]
    [MethodDataSource(nameof(DisplayStringR1C1))]
    public async Task DisplayStringR1C1DisplaysReferenceInR1C1Style(
        RowCol rowCol,
        string expectedString
    )
    {
        await Assert.That(rowCol.GetDisplayStringR1C1()).IsEqualTo(expectedString);
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            yield return
            [
                new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1),
                "A1",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Relative, 14, ReferenceAxisType.Absolute, 28, A1),
                "$AB14",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Absolute, 4, ReferenceAxisType.Relative, 26, A1),
                "Z$4",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Absolute, 264, ReferenceAxisType.Absolute, 3, A1),
                "$C$264",
            ];
        }
    }

    public static IEnumerable<object[]> DisplayStringR1C1
    {
        get
        {
            yield return
            [
                new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, R1C1),
                "R[1]C[1]",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Relative, 105, ReferenceAxisType.None, 0, R1C1),
                "R[105]",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, -7, R1C1),
                "C[-7]",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 1, R1C1),
                "R1C1",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 8, R1C1),
                "C8",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.None, 0, R1C1),
                "R1",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0, R1C1),
                "C",
            ];
            yield return
            [
                new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0, R1C1),
                "R",
            ];
        }
    }
}

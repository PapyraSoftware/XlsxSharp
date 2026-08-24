namespace XlsxSharp.Parser.Tests;

public class ReferenceAreaTests
{
    [Test]
    [MethodDataSource(nameof(DisplayStringA1))]
    public async Task DisplayStringA1DisplaysReferenceInA1Style(
        ReferenceArea reference,
        string expectedString
    )
    {
        await Assert.That(reference.GetDisplayStringA1()).IsEqualTo(expectedString);
        await Assert.That(TokenParser.ParseReference(expectedString, true)).IsEqualTo(reference);
    }

    [Test]
    [MethodDataSource(nameof(DisplayStringR1C1))]
    public async Task DisplayStringR1C1DisplaysReferenceInR1C1Style(
        ReferenceArea reference,
        string expectedString
    )
    {
        await Assert.That(reference.GetDisplayStringR1C1()).IsEqualTo(expectedString);
        await Assert.That(TokenParser.ParseReference(expectedString, false)).IsEqualTo(reference);
    }

    public static IEnumerable<object[]> DisplayStringA1
    {
        get
        {
            // When both corners are same, only one is rendered.
            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1),
                    new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1)
                ),
                "A1",
            ];

            // When both corners are same, only one is rendered.
            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1),
                    new RowCol(ReferenceAxisType.Relative, 5, ReferenceAxisType.Relative, 3, A1)
                ),
                "A1:C5",
            ];

            // Row span
            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 6, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Relative, 6, ReferenceAxisType.None, 0, A1)
                ),
                "6:6",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 6, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Relative, 8, ReferenceAxisType.None, 0, A1)
                ),
                "6:8",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 65, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Absolute, 745, ReferenceAxisType.None, 0, A1)
                ),
                "$65:$745",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 79, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Absolute, 999, ReferenceAxisType.None, 0, A1)
                ),
                "79:$999",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 79, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Relative, 999, ReferenceAxisType.None, 0, A1)
                ),
                "$79:999",
            ];

            // Col span
            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 5, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 5, A1)
                ),
                "E:E",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 2, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 4, A1)
                ),
                "B:D",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 27, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 53, A1)
                ),
                "$AA:$BA",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 96, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 6663, A1)
                ),
                "CR:$IVG",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Absolute, 96, A1),
                    new RowCol(None, 0, Relative, 6663, A1)
                ),
                "$CR:IVG",
            ];
        }
    }

    public static IEnumerable<object[]> DisplayStringR1C1
    {
        get
        {
            yield return
            [
                new ReferenceArea(
                    new RowCol(Absolute, 7, Absolute, 1, R1C1),
                    new RowCol(Absolute, 7, Absolute, 1, R1C1)
                ),
                "R7C1",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(Absolute, 7, Absolute, 1, R1C1),
                    new RowCol(Relative, 0, None, 0, R1C1)
                ),
                "R7C1:R",
            ];

            // Row span
            yield return
            [
                new ReferenceArea(
                    new RowCol(Relative, 6, None, 0, R1C1),
                    new RowCol(Relative, 6, None, 0, R1C1)
                ),
                "R[6]",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(Relative, 6, None, 0, R1C1),
                    new RowCol(Relative, 8, None, 0, R1C1)
                ),
                "R[6]:R[8]",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(Absolute, 65, None, 0, R1C1),
                    new RowCol(Absolute, 745, None, 0, R1C1)
                ),
                "R65:R745",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(Relative, 79, None, 0, R1C1),
                    new RowCol(Absolute, 999, None, 0, R1C1)
                ),
                "R[79]:R999",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(Absolute, 79, None, 0, R1C1),
                    new RowCol(Relative, 999, None, 0, R1C1)
                ),
                "R79:R[999]",
            ];

            // Col span
            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Relative, 2, R1C1),
                    new RowCol(None, 0, Relative, 2, R1C1)
                ),
                "C[2]",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Relative, 2, R1C1),
                    new RowCol(None, 0, Relative, 4, R1C1)
                ),
                "C[2]:C[4]",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Absolute, 27, R1C1),
                    new RowCol(None, 0, Absolute, 53, R1C1)
                ),
                "C27:C53",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Relative, 96, R1C1),
                    new RowCol(None, 0, Absolute, 663, R1C1)
                ),
                "C[96]:C663",
            ];

            yield return
            [
                new ReferenceArea(
                    new RowCol(None, 0, Absolute, 96, R1C1),
                    new RowCol(None, 0, Relative, 663, R1C1)
                ),
                "C96:C[663]",
            ];
        }
    }
}

using XlsxSharp.Parser.Rolex;

namespace XlsxSharp.Parser.Tests.Lexers;

public class R1C1ReferenceTokenTests
{
    [Test]
    [MethodDataSource(nameof(TestDataOneCorner))]
    [MethodDataSource(nameof(TestDataTwoCorners))]
    public async Task ParseExtractsInformationFromToken(
        string token,
        int[] expectedTokens,
        ReferenceArea expectedReference
    )
    {
        await Assert
            .That(RolexLexer.GetTokensR1C1(token).Select(x => x.SymbolId))
            .IsEquivalentTo(expectedTokens.Concat(new[] { Token.EofSymbolId }));
        ReferenceArea reference = TokenParser.ParseReference(token.AsSpan(), false);
        await Assert.That(reference).IsEqualTo(expectedReference);
    }

    public static IEnumerable<object[]> TestDataOneCorner
    {
        get
        {
            // The `C` is a shortcut for `C[0]`
            yield return
            [
                "C",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0, R1C1)
                ),
            ];

            yield return
            [
                "C[-14]",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, -14, R1C1)
                ),
            ];

            yield return
            [
                "C75",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 75, R1C1)
                ),
            ];

            // The `R` is a shortcut for `R[0]`
            yield return
            [
                "R",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0, R1C1)
                ),
            ];

            yield return
            [
                "R[-14]",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, -14, ReferenceAxisType.None, 0, R1C1)
                ),
            ];

            yield return
            [
                "R75",
                new[] { Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 75, ReferenceAxisType.None, 0, R1C1)
                ),
            ];

            yield return
            [
                "RC",
                new[] { Token.A1_CELL },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.Relative, 0, R1C1)
                ),
            ];

            yield return
            [
                "R[7]C2",
                new[] { Token.A1_CELL },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 7, ReferenceAxisType.Absolute, 2, R1C1)
                ),
            ];

            yield return
            [
                "R812C[7]",
                new[] { Token.A1_CELL },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 812, ReferenceAxisType.Relative, 7, R1C1)
                ),
            ];
        }
    }

    public static IEnumerable<object[]> TestDataTwoCorners
    {
        get
        {
            yield return
            [
                "R1C2:R3C4",
                new[] { Token.A1_CELL, Token.COLON, Token.A1_CELL },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 1, ReferenceAxisType.Absolute, 2, R1C1),
                    new RowCol(ReferenceAxisType.Absolute, 3, ReferenceAxisType.Absolute, 4, R1C1)
                ),
            ];

            yield return
            [
                "C:R", // C[0]:R[0], technically legal
                new[] { Token.A1_SPAN_REFERENCE, Token.COLON, Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0, R1C1),
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0, R1C1)
                ),
            ];

            yield return
            [
                "R[-1]C[-2]:R[-3]C[-4]",
                new[] { Token.A1_CELL, Token.COLON, Token.A1_CELL },
                new ReferenceArea(
                    new RowCol(
                        ReferenceAxisType.Relative,
                        -1,
                        ReferenceAxisType.Relative,
                        -2,
                        R1C1
                    ),
                    new RowCol(ReferenceAxisType.Relative, -3, ReferenceAxisType.Relative, -4, R1C1)
                ),
            ];

            yield return
            [
                "R:C", // R[0]:C[0], technically legal
                new[] { Token.A1_SPAN_REFERENCE, Token.COLON, Token.A1_SPAN_REFERENCE },
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 0, ReferenceAxisType.None, 0, R1C1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 0, R1C1)
                ),
            ];
        }
    }
}

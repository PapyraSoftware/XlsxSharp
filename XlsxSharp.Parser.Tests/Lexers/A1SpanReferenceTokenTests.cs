using static XlsxSharp.Parser.ReferenceAxisType;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// Test of a parsing of a token <c>A1_SPAN_REFERENCE</c>.
/// <code>
/// A1_SPAN_REFERENCE
///        : A1_COLUMN ':' A1_COLUMN
///        | A1_ROW ':' A1_ROW
///        ;
/// </code>
/// </summary>
public class A1SpanReferenceTokenTests
{
    [Test]
    public async Task ParseRowRange()
    {
        // Check A1_ROW ':' A1_ROW path
        await AssertAreaReferenceToken(
            "1:1",
            new ReferenceArea(
                new RowCol(Relative, 1, None, 0, ReferenceStyle.A1),
                new RowCol(Relative, 1, None, 0, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "$5:10",
            new ReferenceArea(
                new RowCol(Absolute, 5, None, 0, ReferenceStyle.A1),
                new RowCol(Relative, 10, None, 0, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "7:$3",
            new ReferenceArea(
                new RowCol(Relative, 7, None, 0, ReferenceStyle.A1),
                new RowCol(Absolute, 3, None, 0, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "$1048576:$1048576",
            new ReferenceArea(
                new RowCol(Absolute, RowCol.MaxRow, None, 0, ReferenceStyle.A1),
                new RowCol(Absolute, RowCol.MaxRow, None, 0, ReferenceStyle.A1)
            )
        );
    }

    [Test]
    public async Task ParseColumnRange()
    {
        // Check A1_COLUMN ':' A1_COLUMN path
        await AssertAreaReferenceToken(
            "A:A",
            new ReferenceArea(
                new RowCol(None, 0, Relative, 1, ReferenceStyle.A1),
                new RowCol(None, 0, Relative, 1, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "RW:ST",
            new ReferenceArea(
                new RowCol(None, 0, Relative, 491, ReferenceStyle.A1),
                new RowCol(None, 0, Relative, 514, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "$C:D",
            new ReferenceArea(
                new RowCol(None, 0, Absolute, 3, ReferenceStyle.A1),
                new RowCol(None, 0, Relative, 4, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "E:$C",
            new ReferenceArea(
                new RowCol(None, 0, Relative, 5, ReferenceStyle.A1),
                new RowCol(None, 0, Absolute, 3, ReferenceStyle.A1)
            )
        );
        await AssertAreaReferenceToken(
            "$XFD:$XFD",
            new ReferenceArea(
                new RowCol(None, 0, Absolute, RowCol.MaxCol, ReferenceStyle.A1),
                new RowCol(None, 0, Absolute, RowCol.MaxCol, ReferenceStyle.A1)
            )
        );
    }

    private static async Task AssertAreaReferenceToken(
        string token,
        ReferenceArea expectedReference
    )
    {
        await AssertFormula.AssertTokenType(token, Token.A1_SPAN_REFERENCE);
        ReferenceArea reference = TokenParser.ParseReference(token, true);
        await Assert.That(reference).IsEqualTo(expectedReference);
    }
}

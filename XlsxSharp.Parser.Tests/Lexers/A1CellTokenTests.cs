using static XlsxSharp.Parser.ReferenceAxisType;

namespace XlsxSharp.Parser.Tests.Lexers;

/// <summary>
/// Test of a parsing of a token <c>A1_CELL</c>.
/// <code>
/// A1_CELL
///     : A1_COLUMN A1_ROW
///     ;
/// </code>
/// </summary>
public class A1CellTokenTests
{
    [Test]
    public async Task ParseA1Cell()
    {
        // Check A1_CELL path
        await AssertAreaReferenceToken(
            "$B$3",
            new ReferenceArea(Absolute, 3, Absolute, 2, ReferenceStyle.A1)
        );
        await AssertAreaReferenceToken(
            "A1",
            new ReferenceArea(Relative, 1, Relative, 1, ReferenceStyle.A1)
        );
        await AssertAreaReferenceToken(
            "XFD1",
            new ReferenceArea(Relative, 1, Relative, RowCol.MaxCol, ReferenceStyle.A1)
        );
        await AssertAreaReferenceToken(
            "A1048576",
            new ReferenceArea(Relative, RowCol.MaxRow, Relative, 1, ReferenceStyle.A1)
        );
        await AssertAreaReferenceToken(
            "$XFD$1048576",
            new ReferenceArea(Absolute, RowCol.MaxRow, Absolute, RowCol.MaxCol, ReferenceStyle.A1)
        );
    }

    private static async Task AssertAreaReferenceToken(
        string token,
        ReferenceArea expectedReference
    )
    {
        ReferenceArea reference = TokenParser.ParseReference(token, true);
        await Assert.That(reference).IsEqualTo(expectedReference);
    }
}

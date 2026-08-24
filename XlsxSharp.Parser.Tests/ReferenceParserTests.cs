namespace XlsxSharp.Parser.Tests;

public class ReferenceParserTests
{
    [Test]
    [MethodDataSource(nameof(ParseA1TestCases))]
    public async Task ParseA1ParsesCellAreaOrRowspanOrColspan(
        string text,
        ReferenceArea expectedReference
    )
    {
        await Assert.That(ReferenceParser.ParseA1(text)).IsEqualTo(expectedReference);
    }

    [Test]
    public void ParseA1RequiresArgument()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() => ReferenceParser.ParseA1(null!));
    }

    [Test]
    public void ParseA1ThrowsOnNonReferences()
    {
        Assert.ThrowsExactly<ParsingException>(() => ReferenceParser.ParseA1("HELLO"));
    }

    [Test]
    [MethodDataSource(nameof(ParseA1TestCases))]
    public async Task TryParseA1ParsesCellAreaOrRowspanOrColspan(
        string text,
        ReferenceArea expectedReference
    )
    {
        bool success = ReferenceParser.TryParseA1(text, out ReferenceArea area);
        await Assert.That(success).IsTrue();
        await Assert.That(area).IsEqualTo(expectedReference);
    }

    [Test]
    public void TryParseA1RequiresArgument()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() => ReferenceParser.TryParseA1(null!, out _));
    }

    [Test]
    public async Task ParseA1ReturnsFalseOnNonReferences()
    {
        bool success = ReferenceParser.TryParseA1("HELLO", out ReferenceArea area);
        await Assert.That(success).IsFalse();
        await Assert.That(area).IsEqualTo(default);
    }

    [Test]
    [MethodDataSource(nameof(ParseSheetA1TestCases))]
    public async Task TryParseSheetA1AcceptsAreaOrRowspanOrColspanWithSheet(
        string text,
        string expectedSheet,
        ReferenceArea expectedArea
    )
    {
        bool success = ReferenceParser.TryParseSheetA1(
            text,
            out string sheet,
            out ReferenceArea area
        );
        await Assert.That(success).IsTrue();
        await Assert.That(sheet).IsEqualTo(expectedSheet);
        await Assert.That(area).IsEqualTo(expectedArea);
    }

    [Test]
    public async Task TryParseSheetA1CantParseWorkbookIndex()
    {
        bool success = ReferenceParser.TryParseSheetA1("[1]Sheet!A1", out _, out _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public async Task TryParseSheetA1CantParseReferenceWithoutSheet()
    {
        bool success = ReferenceParser.TryParseSheetA1("A1", out _, out _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public void TryParseSheetA1RequiresArgument()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() =>
            ReferenceParser.TryParseSheetA1(null!, out _, out _)
        );
    }

    [Test]
    [Arguments("Sheet!Name", "Sheet", "Name")]
    [Arguments("'Hello World'!Name", "Hello World", "Name")]
    [Arguments("' John''s World! '!Name", " John's World! ", "Name")]
    public async Task TryParseSheetNameParsesSheetAndName(
        string text,
        string expectedSheet,
        string expectedName
    )
    {
        bool success = ReferenceParser.TryParseSheetName(text, out string sheet, out string name);

        await Assert.That(success).IsTrue();
        await Assert.That(sheet).IsEqualTo(expectedSheet);
        await Assert.That(name).IsEqualTo(expectedName);
    }

    [Test]
    public void TryParseSheetNameRequiresText()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() =>
            ReferenceParser.TryParseSheetName(null!, out _, out _)
        );
    }

    [Test]
    [Arguments("Name")]
    [Arguments("some_name")]
    [Arguments("A1")]
    public async Task TryParseSheetNameCantParsePureName(string text)
    {
        bool success = ReferenceParser.TryParseSheetName(text, out _, out _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    [Arguments("Sheet!Name", "Sheet", "Name")]
    [Arguments("'Hello World'!Data", "Hello World", "Data")]
    [Arguments("' John''s World! '!Quarter1", " John's World! ", "Quarter1")]
    public async Task TryParseNameParsesSheetAndName(
        string text,
        string expectedSheet,
        string expectedName
    )
    {
        bool success = ReferenceParser.TryParseName(text, out string? sheet, out string name);
        await Assert.That(success).IsTrue();
        await Assert.That(sheet).IsEqualTo(expectedSheet);
        await Assert.That(name).IsEqualTo(expectedName);
    }

    [Test]
    public void TryParseNameRequiresText()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() =>
            ReferenceParser.TryParseName(null!, out _, out _)
        );
    }

    [Test]
    [Arguments("Name")]
    [Arguments("some_name")]
    public async Task TryParseNameParsesName(string text)
    {
        bool success = ReferenceParser.TryParseName(text, out string? sheet, out string name);
        await Assert.That(success).IsTrue();
        Assert.Null(sheet);
        await Assert.That(name).IsEqualTo(text);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("$BC$1")]
    [Arguments("Sheet!A1")]
    [Arguments("14")]
    [Arguments("\"Text\"")]
    public async Task TryParseNameCantParseAnythingButName(string text)
    {
        bool success = ReferenceParser.TryParseName(text, out _, out _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    [MethodDataSource(nameof(ParseA1TestCases))]
    public async Task TryParseA1UnifiedCanParseLocalReference(
        string text,
        ReferenceArea expectedReference
    )
    {
        bool success = ReferenceParser.TryParseA1(
            text,
            out string? sheet,
            out ReferenceArea reference
        );
        await Assert.That(success).IsTrue();
        Assert.Null(sheet);
        await Assert.That(reference).IsEqualTo(expectedReference);
    }

    [Test]
    [MethodDataSource(nameof(ParseSheetA1TestCases))]
    public async Task TryParseA1UnifiedCanParseSheetReference(
        string text,
        string expectedSheet,
        ReferenceArea expectedReference
    )
    {
        bool success = ReferenceParser.TryParseA1(
            text,
            out string? sheet,
            out ReferenceArea reference
        );
        await Assert.That(success).IsTrue();
        await Assert.That(sheet).IsEqualTo(expectedSheet);
        await Assert.That(reference).IsEqualTo(expectedReference);
    }

    [Test]
    [Arguments("Sheet!Name")]
    [Arguments("Name")]
    [Arguments("1")]
    public async Task TryParseA1UnifiedCantParseAnythingButReference(string text)
    {
        bool success = ReferenceParser.TryParseA1(text, out _, out _);
        await Assert.That(success).IsFalse();
    }

    [Test]
    public void TryParseA1UnifiedRequiresArgument()
    {
        Assert.ThrowsExactly<ArgumentNullException>(() =>
            ReferenceParser.TryParseA1(null!, out _, out _)
        );
    }

    public static IEnumerable<object[]> ParseSheetA1TestCases
    {
        get
        {
            yield return
            [
                "Sheet!$C$2",
                "Sheet",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 2, ReferenceAxisType.Absolute, 3, A1)
                ),
            ];
            yield return
            [
                "' ''John''s'' Shop! '!C2",
                " 'John's' Shop! ",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 2, ReferenceAxisType.Relative, 3, A1)
                ),
            ];
            yield return
            [
                "Sheet!A1:B2",
                "Sheet",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 1, ReferenceAxisType.Relative, 1, A1),
                    new RowCol(ReferenceAxisType.Relative, 2, ReferenceAxisType.Relative, 2, A1)
                ),
            ];
            yield return
            [
                "'Some Sheet'!C:D",
                "Some Sheet",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 3, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 4, A1)
                ),
            ];
            yield return
            [
                "'!!WARN'!10:$15",
                "!!WARN",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 10, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Absolute, 15, ReferenceAxisType.None, 0, A1)
                ),
            ];
        }
    }

    public static IEnumerable<object[]> ParseA1TestCases
    {
        get
        {
            yield return
            [
                "$C$2",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 2, ReferenceAxisType.Absolute, 3, A1)
                ),
            ];
            yield return
            [
                "AB123",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 123, ReferenceAxisType.Relative, 28, A1)
                ),
            ];
            yield return
            [
                "$C$2:E7",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Absolute, 2, ReferenceAxisType.Absolute, 3, A1),
                    new RowCol(ReferenceAxisType.Relative, 7, ReferenceAxisType.Relative, 5, A1)
                ),
            ];
            yield return
            [
                "$C:F",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Absolute, 3, A1),
                    new RowCol(ReferenceAxisType.None, 0, ReferenceAxisType.Relative, 6, A1)
                ),
            ];
            yield return
            [
                "10:$15",
                new ReferenceArea(
                    new RowCol(ReferenceAxisType.Relative, 10, ReferenceAxisType.None, 0, A1),
                    new RowCol(ReferenceAxisType.Absolute, 15, ReferenceAxisType.None, 0, A1)
                ),
            ];
        }
    }
}

namespace XlsxSharp.Parser.Tests.Lexers;

// Tests of parsing SHEET_RANGE_PREFIX
public class SheetRangePrefixTokenTests
{
    [Test]
    [MethodDataSource(nameof(Data))]
    public async Task TokenDataAreExtractedAndUnescaped(
        string tokenText,
        int? expectedWorkbookIndex,
        string expectedFirstSheetName,
        string expectedSecondSheetName
    )
    {
        TokenParser.ParseSheetRangePrefix(
            tokenText,
            out int? workbookIndex,
            out string firstSheetName,
            out string secondSheetName
        );

        await Assert.That(workbookIndex).IsEqualTo(expectedWorkbookIndex);
        await Assert.That(firstSheetName).IsEqualTo(expectedFirstSheetName);
        await Assert.That(secondSheetName).IsEqualTo(expectedSecondSheetName);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            // Special case for sheet range starting with column
            yield return ["JAN:DEC!", null, "JAN", "DEC"];
            yield return ["PGG:EGAS_Ele!", null, "PGG", "EGAS_Ele"];

            yield return ["[1]first:second!", 1, "first", "second"];

            // No escape, but enclosed in tick
            yield return ["'[1]first:second'!", 1, "first", "second"];
            yield return ["'first:second'!", null, "first", "second"];

            // Test correct escaping
            yield return ["'Monty''s:Johnny''s'!", null, "Monty's", "Johnny's"];

            // multiple escapes
            yield return ["'[7]a''''''b:c''''d'!", 7, "a'''b", "c''d"];

            // single character name
            yield return ["'[6]a:b'!", 6, "a", "b"];
        }
    }
}

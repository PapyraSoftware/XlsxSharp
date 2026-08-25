namespace XlsxSharp.Parser.Tests.Lexers;

public class SingleSheetPrefixTokenTests
{
    [Test]
    [MethodDataSource(nameof(Data))]
    public async Task TokenDataAreExtractedAndUnescaped(
        string tokenText,
        int? expectedWorkbookIndex,
        string expectedSheetName
    )
    {
        TokenParser.ParseSingleSheetPrefix(tokenText, out int? workbookIndex, out string sheetName);

        await Assert.That(workbookIndex).IsEqualTo(expectedWorkbookIndex);
        await Assert.That(sheetName).IsEqualTo(expectedSheetName);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            yield return ["sheet!", null, "sheet"];
            yield return ["[7]sheet!", 7, "sheet"];
            yield return ["'sheet name'!", null, "sheet name"];
            yield return ["'[2]Monty''s'!", 2, "Monty's"];
            yield return ["'[25]a''''''b'!", 25, "a'''b"];
        }
    }
}

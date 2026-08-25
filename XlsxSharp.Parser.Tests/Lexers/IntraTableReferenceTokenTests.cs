namespace XlsxSharp.Parser.Tests.Lexers;

public class IntraTableReferenceTokenTests
{
    [Test]
    [MethodDataSource(nameof(Data))]
    public async Task TokenDataAreExtractedAndUnescaped(
        string tokenText,
        StructuredReferenceArea expectedArea,
        string expectedFirstColumn,
        string expectedLastColumn
    )
    {
        TokenParser.ParseIntraTableReference(
            tokenText,
            out StructuredReferenceArea area,
            out string? firstColumn,
            out string? lastColumn
        );

        await Assert.That(area).IsEqualTo(expectedArea);
        await Assert.That(firstColumn).IsEqualTo(expectedFirstColumn);
        await Assert.That(lastColumn).IsEqualTo(expectedLastColumn);
    }

    public static IEnumerable<object?[]> Data
    {
        get
        {
            // Portions area
            // INTRA_TABLE_REFERENCE : KEYWORD
            yield return ["[#All]", StructuredReferenceArea.All, null, null];
            yield return ["[#Data]", StructuredReferenceArea.Data, null, null];
            yield return ["[#Headers]", StructuredReferenceArea.Headers, null, null];
            yield return ["[#Totals]", StructuredReferenceArea.Totals, null, null];
            yield return ["[#This Row]", StructuredReferenceArea.ThisRow, null, null];

            // Empty simple column, per grammar, the SIMPLE_COLUMN_NAME is optional
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']'
            yield return ["[]", StructuredReferenceArea.None, null, null];

            // Simple column
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']'
            yield return ["[Col]", StructuredReferenceArea.None, "Col", null];
            yield return
            [
                "[Name with space]",
                StructuredReferenceArea.None,
                "Name with space",
                null,
            ];

            // Escaped characters
            // INTRA_TABLE_REFERENCE : '[' SIMPLE_COLUMN_NAME? ']'
            // where column name is a possible value of a ESCAPE_COLUMN_CHARACTER
            yield return ["['[']]", StructuredReferenceArea.None, "[]", null];
            yield return ["['''#]", StructuredReferenceArea.None, "'#", null];
            yield return ["['[']'''#]", StructuredReferenceArea.None, "[]'#", null];

            // INTRA_TABLE_REFERENCE : SPACED_LBRACKET INNER_REFERENCE SPACED_RBRACKET
            // where inner reference is `COLUMN_RANGE : COLUMN(':' COLUMN)?`
            yield return ["[[First]]", StructuredReferenceArea.None, "First", null];
            yield return ["[[First]:[Last]]", StructuredReferenceArea.None, "First", "Last"];
            yield return ["[[First]:Last]", StructuredReferenceArea.None, "First", "Last"];
            yield return ["[First:[Last]]", StructuredReferenceArea.None, "First", "Last"];
            yield return ["[First:Last]", StructuredReferenceArea.None, "First", "Last"];

            // fragment INNER_REFERENCE : KEYWORD_LIST SPACED_COMMA COLUMN_RANGE
            // where KEYWORD_LIST is just a KEYWORD
            yield return ["[[#All],[First]]", StructuredReferenceArea.All, "First", null];
            yield return
            [
                "[[#Data],[First]:[Last]]",
                StructuredReferenceArea.Data,
                "First",
                "Last",
            ];
            yield return
            [
                "[[#Headers],[First]:Last]",
                StructuredReferenceArea.Headers,
                "First",
                "Last",
            ];
            yield return
            [
                "[[#Totals],First:[Last]]",
                StructuredReferenceArea.Totals,
                "First",
                "Last",
            ];
            yield return
            [
                "[[#This Row],First:Last]",
                StructuredReferenceArea.ThisRow,
                "First",
                "Last",
            ];

            // fragment INNER_REFERENCE : KEYWORD_LIST SPACED_COMMA COLUMN_RANGE
            // where KEYWORD_LIST | '[#Headers]' SPACED_COMMA '[#Data]' | '[#Data]' SPACED_COMMA '[#Totals]'
            yield return
            [
                "[[#Headers],[#Data],[Col]]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "Col",
                null,
            ];
            yield return
            [
                "[[#Headers],[#Data],[First col]:[Last col]]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "First col",
                "Last col",
            ];
            yield return
            [
                "[[#Headers],[#Data],First:Last]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "First",
                "Last",
            ];
            yield return
            [
                "[[#Headers],[#Data],[First]:Last]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "First",
                "Last",
            ];
            yield return
            [
                "[[#Headers],[#Data],First:[Last]]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "First",
                "Last",
            ];

            // spaces are ignored
            yield return
            [
                "[  [#Headers]  ,  [#Data]  ,  [First col]:[Last col]  ]",
                StructuredReferenceArea.Headers | StructuredReferenceArea.Data,
                "First col",
                "Last col",
            ];
        }
    }
}

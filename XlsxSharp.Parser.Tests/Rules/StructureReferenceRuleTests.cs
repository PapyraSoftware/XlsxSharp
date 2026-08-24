namespace XlsxSharp.Parser.Tests.Rules;

public class StructureReferenceRuleTests
{
    [Test]
    [MethodDataSource(nameof(TestCases))]
    public async Task StructureReferenceIsParsedToANode(string formula, AstNode expectedNode)
    {
        await AssertFormula.SingleNodeParsed(formula, expectedNode);
    }

    public static IEnumerable<object[]> TestCases
    {
        get
        {
            // structure_reference : INTRA_TABLE_REFERENCE
            yield return
            [
                "[Column]",
                new StructureReferenceNode(null, StructuredReferenceArea.None, "Column", "Column"),
            ];

            yield return
            [
                "[#Totals]",
                new StructureReferenceNode(null, StructuredReferenceArea.Totals, null, null),
            ];

            yield return
            [
                "[]",
                new StructureReferenceNode(null, StructuredReferenceArea.None, null, null),
            ];

            yield return
            [
                "[[#Data],[First Column]:[Last Column]]",
                new StructureReferenceNode(
                    null,
                    StructuredReferenceArea.Data,
                    "First Column",
                    "Last Column"
                ),
            ];

            // structure_reference : NAME INTRA_TABLE_REFERENCE
            yield return
            [
                "SomeTable[Column]",
                new StructureReferenceNode(
                    "SomeTable",
                    StructuredReferenceArea.None,
                    "Column",
                    "Column"
                ),
            ];

            // structure_reference: BOOK_PREFIX NAME INTRA_TABLE_REFERENCE
            yield return
            [
                "[4]!SomeTable[Column]",
                new ExternalStructureReferenceNode(
                    4,
                    "SomeTable",
                    StructuredReferenceArea.None,
                    "Column",
                    "Column"
                ),
            ];
        }
    }
}

namespace XlsxSharp.Parser.Tests.Lexers;

public class CellFunctionListTokenTests
{
    [Test]
    public async Task IgnoresTrailingWhitespaces()
    {
        RowCol expected = new(1, 1, A1);
        await Assert.That(TokenParser.ExtractCellFunction("A1(  ")).IsEqualTo(expected);
    }

    [Test]
    [MethodDataSource(nameof(TestData))]
    public async Task AcceptsAbsoluteAndRelativeCellAddresses(string token, RowCol expectedCell)
    {
        await Assert.That(TokenParser.ExtractCellFunction(token)).IsEqualTo(expectedCell);
    }

    public static IEnumerable<object[]> TestData
    {
        get
        {
            yield return ["A1(", new RowCol(1, 1, A1)];
            yield return ["$A$1(", new RowCol(true, 1, true, 1, A1)];
            yield return ["$B3(", new RowCol(false, 3, true, 2, A1)];
            yield return ["B$3(", new RowCol(true, 3, false, 2, A1)];
        }
    }
}

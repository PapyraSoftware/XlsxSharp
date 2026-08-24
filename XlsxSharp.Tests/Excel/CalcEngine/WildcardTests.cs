using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class WildcardTests
{
    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public void EmptyPatternMatchesAnyString(string text) =>
        ClassicAssert.AreEqual(0, SearchWildcard(text, string.Empty));

    [Test]
    [Arguments("", "abc", 0)]
    [Arguments("a", "abc", 0)]
    [Arguments("ab", "abc", 0)]
    [Arguments("abc", "abc", 0)]
    [Arguments("bc", "abc", 1)]
    [Arguments("c", "abc", 2)]
    public void SubstringOfTextMatchesText(
        string substringPattern,
        string text,
        int expectedIndex
    ) => ClassicAssert.AreEqual(expectedIndex, SearchWildcard(text, substringPattern));

    [Test]
    [Arguments("abcd", "abc")]
    public void PatternNotInTextReturnsNegativeOne(string pattern, string text) =>
        ClassicAssert.AreEqual(-1, SearchWildcard(text, pattern));

    [Test]
    public void PatternComparisonIsCaseInsensitive() =>
        ClassicAssert.AreEqual(1, SearchWildcard("zabcd", "AbCd"));

    [Test]
    public void TildeIsEscapeChar() => ClassicAssert.AreEqual(1, SearchWildcard("_abc_", "~a~B~c"));

    [Test]
    [Arguments("~*", "*", 0)]
    [Arguments("~*", "a", -1)]
    [Arguments("~?", "?", 0)]
    [Arguments("~?", "a", -1)]
    [Arguments("~a~b~", "ab", 0)]
    public void EscapedWildcardsAreMatchedAsChars(
        string pattern,
        string text,
        int expectedPosition
    ) => ClassicAssert.AreEqual(expectedPosition, SearchWildcard(text, pattern));

    [Test]
    public void QuestionMarkWildcardMatchesAnyChar() =>
        ClassicAssert.AreEqual(0, SearchWildcard("abc", "a?c"));

    [Test]
    [Arguments("abcd", "ab*cd", 0)]
    [Arguments(@"aaab_____cd", "ab*cd", 2)]
    [Arguments("*abc*", "***a*b*c***", 0)]
    public void StarWildcardMatchesAnyNumberOfChars(string text, string pattern, int index) =>
        ClassicAssert.AreEqual(index, SearchWildcard(text, pattern));

    [Test]
    public void UnpairedEscapeCharAtTheEndOfPatternIsNotChar() =>
        ClassicAssert.AreEqual(0, SearchWildcard("a", "a~"));

    [Test]
    public void StarWildcardAtTheBeginningMatchesFirstChar() =>
        ClassicAssert.AreEqual(0, SearchWildcard("abcccd", "*ccd"));

    [Test]
    public void PatternSizeIsLimitedTo255Chars()
    {
        ClassicAssert.AreEqual(0, SearchWildcard(new string('a', 1000), new string('a', 255)));

        ClassicAssert.AreEqual(-1, SearchWildcard(new string('a', 1000), new string('a', 256)));
    }

    [Test]
    [Arguments("?", "a", true)]
    [Arguments("?", "ab", false)]
    [Arguments("a?", "ab", true)]
    [Arguments("a?", "abc", false)]
    [Arguments("?b", "ab", true)]
    [Arguments("?b", "aab", false)]
    [Arguments("a*", "abc", true)]
    [Arguments("*a*", "abc", true)]
    [Arguments("*c", "abc", true)]
    [Arguments("*a*a", "abc", false)]
    [Arguments("*a*a", "aba", true)]
    [Arguments("*a*a", @"zaba", true)]
    [Arguments("a*", @"zaba", false)]
    public void Matches(string pattern, string text, bool matches) =>
        ClassicAssert.AreEqual(matches, new Wildcard(pattern).Matches(text.AsSpan()));

    private static int SearchWildcard(string text, string pattern) =>
        new Wildcard(pattern).Search(text.AsSpan());
}

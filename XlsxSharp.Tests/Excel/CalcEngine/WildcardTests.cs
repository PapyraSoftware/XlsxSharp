using System;
using NUnit.Framework;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

[TestFixture]
public class WildcardTests
{
    [TestCase("")]
    [TestCase("abc")]
    public void EmptyPatternMatchesAnyString(string text) =>
        Assert.AreEqual(0, SearchWildcard(text, string.Empty));

    [TestCase("", "abc", 0)]
    [TestCase("a", "abc", 0)]
    [TestCase("ab", "abc", 0)]
    [TestCase("abc", "abc", 0)]
    [TestCase("bc", "abc", 1)]
    [TestCase("c", "abc", 2)]
    public void SubstringOfTextMatchesText(
        string substringPattern,
        string text,
        int expectedIndex
    ) => Assert.AreEqual(expectedIndex, SearchWildcard(text, substringPattern));

    [TestCase("abcd", "abc")]
    public void PatternNotInTextReturnsNegativeOne(string pattern, string text) =>
        Assert.AreEqual(-1, SearchWildcard(text, pattern));

    [Test]
    public void PatternComparisonIsCaseInsensitive() =>
        Assert.AreEqual(1, SearchWildcard("zabcd", "AbCd"));

    [Test]
    public void TildeIsEscapeChar() => Assert.AreEqual(1, SearchWildcard("_abc_", "~a~B~c"));

    [TestCase("~*", "*", 0)]
    [TestCase("~*", "a", -1)]
    [TestCase("~?", "?", 0)]
    [TestCase("~?", "a", -1)]
    [TestCase("~a~b~", "ab", 0)]
    public void EscapedWildcardsAreMatchedAsChars(
        string pattern,
        string text,
        int expectedPosition
    ) => Assert.AreEqual(expectedPosition, SearchWildcard(text, pattern));

    [Test]
    public void QuestionMarkWildcardMatchesAnyChar() =>
        Assert.AreEqual(0, SearchWildcard("abc", "a?c"));

    [TestCase("abcd", "ab*cd", 0)]
    [TestCase(@"aaab_____cd", "ab*cd", 2)]
    [TestCase("*abc*", "***a*b*c***", 0)]
    public void StarWildcardMatchesAnyNumberOfChars(string text, string pattern, int index) =>
        Assert.AreEqual(index, SearchWildcard(text, pattern));

    [Test]
    public void UnpairedEscapeCharAtTheEndOfPatternIsNotChar() =>
        Assert.AreEqual(0, SearchWildcard("a", "a~"));

    [Test]
    public void StarWildcardAtTheBeginningMatchesFirstChar() =>
        Assert.AreEqual(0, SearchWildcard("abcccd", "*ccd"));

    [Test]
    public void PatternSizeIsLimitedTo255Chars()
    {
        Assert.AreEqual(0, SearchWildcard(new string('a', 1000), new string('a', 255)));

        Assert.AreEqual(-1, SearchWildcard(new string('a', 1000), new string('a', 256)));
    }

    [TestCase("?", "a", true)]
    [TestCase("?", "ab", false)]
    [TestCase("a?", "ab", true)]
    [TestCase("a?", "abc", false)]
    [TestCase("?b", "ab", true)]
    [TestCase("?b", "aab", false)]
    [TestCase("a*", "abc", true)]
    [TestCase("*a*", "abc", true)]
    [TestCase("*c", "abc", true)]
    [TestCase("*a*a", "abc", false)]
    [TestCase("*a*a", "aba", true)]
    [TestCase("*a*a", @"zaba", true)]
    [TestCase("a*", @"zaba", false)]
    public void Matches(string pattern, string text, bool matches) =>
        Assert.AreEqual(matches, new Wildcard(pattern).Matches(text.AsSpan()));

    private static int SearchWildcard(string text, string pattern) =>
        new Wildcard(pattern).Search(text.AsSpan());
}

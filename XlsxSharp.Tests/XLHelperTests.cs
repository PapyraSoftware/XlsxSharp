using System;

namespace XlsxSharp.Tests;

public class XlHelperTests
{
    [Test]
    public void IsValidColumnTest()
    {
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn(""));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("1"));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("A1"));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("AA1"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("A"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("AA"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("AAA"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("Z"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("ZZ"));
        ClassicAssert.AreEqual(true, XLHelper.IsValidColumn("XFD"));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("ZAA"));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("XZA"));
        ClassicAssert.AreEqual(false, XLHelper.IsValidColumn("XFZ"));
    }

    [Test]
    public void ReplaceRelative1()
    {
        string result = XLHelper.ReplaceRelative("A1", 2, "B");
        ClassicAssert.AreEqual("B2", result);
    }

    [Test]
    public void ReplaceRelative2()
    {
        string result = XLHelper.ReplaceRelative("$A1", 2, "B");
        ClassicAssert.AreEqual("$A2", result);
    }

    [Test]
    public void ReplaceRelative3()
    {
        string result = XLHelper.ReplaceRelative("A$1", 2, "B");
        ClassicAssert.AreEqual("B$1", result);
    }

    [Test]
    public void ReplaceRelative4()
    {
        string result = XLHelper.ReplaceRelative("$A$1", 2, "B");
        ClassicAssert.AreEqual("$A$1", result);
    }

    [Test]
    public void ReplaceRelative5()
    {
        string result = XLHelper.ReplaceRelative("1:1", 2, "B");
        ClassicAssert.AreEqual("2:2", result);
    }

    [Test]
    public void ReplaceRelative6()
    {
        string result = XLHelper.ReplaceRelative("$1:1", 2, "B");
        ClassicAssert.AreEqual("$1:2", result);
    }

    [Test]
    public void ReplaceRelative7()
    {
        string result = XLHelper.ReplaceRelative("1:$1", 2, "B");
        ClassicAssert.AreEqual("2:$1", result);
    }

    [Test]
    public void ReplaceRelative8()
    {
        string result = XLHelper.ReplaceRelative("$1:$1", 2, "B");
        ClassicAssert.AreEqual("$1:$1", result);
    }

    [Test]
    public void ReplaceRelative9()
    {
        string result = XLHelper.ReplaceRelative("A:A", 2, "B");
        ClassicAssert.AreEqual("B:B", result);
    }

    [Test]
    public void ReplaceRelativeA()
    {
        string result = XLHelper.ReplaceRelative("$A:A", 2, "B");
        ClassicAssert.AreEqual("$A:B", result);
    }

    [Test]
    public void ReplaceRelativeB()
    {
        string result = XLHelper.ReplaceRelative("A:$A", 2, "B");
        ClassicAssert.AreEqual("B:$A", result);
    }

    [Test]
    public void ReplaceRelativeC()
    {
        string result = XLHelper.ReplaceRelative("$A:$A", 2, "B");
        ClassicAssert.AreEqual("$A:$A", result);
    }

    [Test]
    [Arguments("Sheet1", "Sheet1")]
    [Arguments("O'Brien's sales", "O'Brien's sales")]
    [Arguments(" data # ", " data # ")]
    [Arguments("data $1.00", "data $1.00")]
    [Arguments("data ", "data?")]
    [Arguments("abc def", "abc/def")]
    [Arguments("data 0 ", "data[0]")]
    [Arguments("data ", "data*")]
    [Arguments("abc def", "abc\\def")]
    [Arguments(" data", "'data")]
    [Arguments("data ", "data'")]
    [Arguments("d'at'a", "d'at'a")]
    [Arguments("sheet a4", "sheet:a4")]
    [Arguments("null", null)]
    [Arguments("empty", "")]
    [Arguments("1234567890123456789012345678901", "1234567890123456789012345678901TOOLONG")]
    public void CreateSafeSheetNames(string expected, string input)
    {
        string actual = XLHelper.CreateSafeSheetName(input);
        ClassicAssert.AreEqual(expected, actual);
    }

    [Test]
    [Arguments("Sheet1", "Sheet1")]
    [Arguments("O'Brien's sales", "O'Brien's sales")]
    [Arguments(" data # ", " data # ")]
    [Arguments("data $1.00", "data $1.00")]
    [Arguments("data?", "data_")]
    [Arguments("abc/def", "abc_def")]
    [Arguments("data[0]", "data_0_")]
    [Arguments("data*", "data_")]
    [Arguments("abc\\def", "abc_def")]
    [Arguments("'data", "_data")]
    [Arguments("data'", "data_")]
    [Arguments("d'at'a", "d'at'a")]
    [Arguments("sheet:a4", "sheet_a4")]
    [Arguments(null, "null")]
    [Arguments("", "empty")]
    [Arguments("1234567890123456789012345678901TOOLONG", "1234567890123456789012345678901")]
    public void CreateSafeSheetNamesWithUnderscore(string input, string expected)
    {
        ClassicAssert.AreEqual(expected, XLHelper.CreateSafeSheetName(input, replaceChar: '_'));
    }

    [Test]
    public void CreateSafeSheetNamesInvalidReplacementChar() =>
        ClassicAssert.Throws<ArgumentException>(() =>
            XLHelper.CreateSafeSheetName("abc\\def", replaceChar: ':')
        );
}

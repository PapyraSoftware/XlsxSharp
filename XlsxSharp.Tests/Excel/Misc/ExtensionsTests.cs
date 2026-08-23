using System;
using DocumentFormat.OpenXml;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Excel.Misc;

public class ExtensionsTests
{
    [Test]
    public void FixNewLines()
    {
        ClassicAssert.AreEqual("\n".FixNewLines(), Environment.NewLine);
        ClassicAssert.AreEqual("\r\n".FixNewLines(), Environment.NewLine);
        ClassicAssert.AreEqual("\rS\n".FixNewLines(), "\rS" + Environment.NewLine);
        ClassicAssert.AreEqual("\r\n\n".FixNewLines(), Environment.NewLine + Environment.NewLine);
    }

    [Test]
    public void DoubleSaveRound()
    {
        double value = 1234.1234567;
        ClassicAssert.AreEqual(value.SaveRound(), Math.Round(value, 6));
    }

    [Test]
    public void DoubleValueSaveRound()
    {
        double value = 1234.1234567;
        ClassicAssert.AreEqual(new DoubleValue(value).SaveRound().Value, Math.Round(value, 6));
    }

    [Test]
    [Arguments("NoEscaping", "NoEscaping")]
    [Arguments("1", "'1'")]
    [Arguments("AB-CD", "'AB-CD'")]
    [Arguments(" AB", "' AB'")]
    [Arguments("Test sheet", "'Test sheet'")]
    [Arguments("O'Kelly", "'O''Kelly'")]
    [Arguments("A2+3", "'A2+3'")]
    [Arguments("A\"B", "'A\"B'")]
    [Arguments("A!B", "'A!B'")]
    [Arguments("A~B", "'A~B'")]
    [Arguments("A^B", "'A^B'")]
    [Arguments("A&B", "'A&B'")]
    [Arguments("A>B", "'A>B'")]
    [Arguments("A<B", "'A<B'")]
    [Arguments("A.B", "A.B")]
    [Arguments(".", "'.'")]
    [Arguments("A_B", "A_B")]
    [Arguments("_", "_")]
    [Arguments("=", "'='")]
    [Arguments("A,B", "'A,B'")]
    [Arguments("A@B", "'A@B'")]
    [Arguments("(Test)", "'(Test)'")]
    [Arguments("A#", "'A#'")]
    [Arguments("A$", "'A$'")]
    [Arguments("A%", "'A%'")]
    [Arguments("ABC1", "'ABC1'")]
    [Arguments("ABCD1", "ABCD1")]
    [Arguments("R1C1", "'R1C1'")]
    [Arguments("A{", "'A{'")]
    [Arguments("A}", "'A}'")]
    [Arguments("A`", "'A`'")]
    [Arguments("Русский", "Русский")]
    [Arguments("日本語", "日本語")]
    [Arguments("한국어", "한국어")]
    [Arguments("Slovenščina", "Slovenščina")]
    [Arguments("", "")]
    [Arguments(null, null)]
    public void CanEscapeSheetName(string sheetName, string expected)
    {
        ClassicAssert.AreEqual(expected, StringExtensions.EscapeSheetName(sheetName));
    }

    [Test]
    [Arguments("TestSheet", "TestSheet")]
    [Arguments("'Test sheet'", "Test sheet")]
    [Arguments("'O''Kelly'", "O'Kelly")]
    public void CanUnescapeSheetName(string sheetName, string expected)
    {
        ClassicAssert.AreEqual(expected, StringExtensions.UnescapeSheetName(sheetName));
    }
}

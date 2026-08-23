using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal class XStringConvertTests
{
    [Test]
    [Arguments("", "")]
    [Arguments("_x000D_", "\r")]
    [Arguments("_x30ab_", "カ")] // Hexadecimal numbers are case insensitive
    [Arguments("_x0009_", "\t")]
    [Arguments("__x0041__", "_A_")]
    [Arguments("A_x0042_C", "ABC")]
    [Arguments("_X0041_", "_X0041_")] // Must be lowercase x in the pattern
    [Arguments("_x263A_", "\u263a")] // Smiley face
    [Arguments("_xD83D__xDE43_", "\ud83d\ude43")] // Astral planes - Upside down smiley face
    [Arguments("Result:_x0009_ _x0057_", "Result:\t W")]
    [Arguments("DE_x005F_xAB50_0161_title", "DE_xAB50_0161_title")]
    [Arguments("_x0001_ _x0002_ _x0003_ _x0004_", "\u0001 \u0002 \u0003 \u0004")]
    [Arguments("_x0005_ _x0006_ _x0007_ _x0008_", "\u0005 \u0006 \u0007 \u0008")]
    [Arguments("_xaaBB_ _xAAbb_", "\uAABB \uAABB")]
    [Arguments(@"_Xceed_Something", @"_Xceed_Something")] // https://github.com/XlsxSharp/XlsxSharp/issues/1154
    [Arguments("_xD83DDE43_", "_xD83DDE43_")] // 8 hex digit name, decoded by XmlConvert.DecodeName, but not by XString
    [Arguments("DE_XAB500161_seo_title", "DE_XAB500161_seo_title")] // https://github.com/XlsxSharp/XlsxSharp/issues/2610
    public void Decodes_encoded_unicode_characters(string sourceText, string expectedText)
    {
        string decodedText = XStringConvert.Decode(sourceText);

        ClassicAssert.AreEqual(expectedText, decodedText);
    }
}

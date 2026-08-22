using System.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.RichText;

namespace XlsxSharp.Tests.Excel.RichText;

[TestFixture]
public class XLImmutableRichTextTests
{
    [Test]
    public void Equals_compares_text_runs_phonetic_runs_and_properties()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        XLRichText richText = (XLRichText)ws.Cell("A1").CreateRichText();
        richText
            .AddText("こんにち")
            .SetBold(true) // Hello in hiragana
            .AddText("は,")
            .SetBold(false) // object marker
            .AddText("世界")
            .SetFontSize(15); // world in kanji
        richText.Phonetics.SetAlignment(XLPhoneticAlignment.Distributed).Add(@"konnichi wa", 0, 6); // world in hiragana

        // Assert equal
        XLImmutableRichText immutableRichText = XLImmutableRichText.Create(richText);
        XLImmutableRichText equalImmutableRichText = XLImmutableRichText.Create(richText);
        Assert.AreEqual(immutableRichText, equalImmutableRichText);

        // Different font of a first run
        richText.ElementAt(0).SetBold(false);
        XLImmutableRichText withDifferentTextRunFont = XLImmutableRichText.Create(richText);
        Assert.AreNotEqual(immutableRichText, withDifferentTextRunFont);
        richText.ElementAt(0).SetBold(true);

        // Different phonetic properties
        richText.Phonetics.SetAlignment(XLPhoneticAlignment.Left);
        XLImmutableRichText withDifferentPhoneticsProps = XLImmutableRichText.Create(richText);
        Assert.AreNotEqual(immutableRichText, withDifferentPhoneticsProps);
        richText.Phonetics.SetAlignment(XLPhoneticAlignment.Distributed);

        // Different phonetic runs
        richText.Phonetics.Add("せかい", 6, 8);
        XLImmutableRichText withDifferentTextPhonetics = XLImmutableRichText.Create(richText);
        Assert.AreNotEqual(immutableRichText, withDifferentTextPhonetics);
    }
}

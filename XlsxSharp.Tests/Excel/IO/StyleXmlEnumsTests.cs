using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel;
using XlsxSharp.Excel.IO;

namespace XlsxSharp.Tests.Excel.IO;

/// <summary>
/// Checks the font enumerations against the SDK conversions they replace, in both directions.
/// </summary>
public class StyleXmlEnumsTests
{
    [Test]
    public void UnderlineMatchesTheSdk() =>
        AssertParity(
            (UnderlineValues v) => v.ToXlsxSharp(),
            (XLFontUnderlineValues v) => v.ToOpenXml(),
            StyleXmlEnums.ParseUnderline,
            v => v.ToXml()
        );

    [Test]
    public void VerticalTextAlignmentMatchesTheSdk() =>
        AssertParity(
            (VerticalAlignmentRunValues v) => v.ToXlsxSharp(),
            (XLFontVerticalTextAlignmentValues v) => v.ToOpenXml(),
            StyleXmlEnums.ParseVerticalTextAlignment,
            v => v.ToXml()
        );

    [Test]
    public void FontSchemeMatchesTheSdk() =>
        AssertParity(
            (FontSchemeValues v) => v.ToXlsxSharp(),
            (XLFontScheme v) => v.ToOpenXmlEnum(),
            StyleXmlEnums.ParseFontScheme,
            v => v.ToXml()
        );

    [Test]
    public void AnUnknownValueIsRejected() =>
        ClassicAssert.Throws<XlsxSharp.IO.PartStructureException>(() =>
            StyleXmlEnums.ParseUnderline("wavy")
        );

    private static void AssertParity<TSdk, TXl>(
        Func<TSdk, TXl> convertWithSdk,
        Func<TXl, TSdk> writeWithSdk,
        Func<string, TXl> parse,
        Func<TXl, string> toXml
    )
        where TSdk : struct, IEnumValue
        where TXl : struct, Enum
    {
        int checkedValues = 0;

        foreach (
            PropertyInfo property in typeof(TSdk).GetProperties(
                BindingFlags.Public | BindingFlags.Static
            )
        )
        {
            if (property.PropertyType != typeof(TSdk))
            {
                continue;
            }

            TSdk value = (TSdk)property.GetValue(null)!;
            string xml = ((IEnumValue)value).Value;

            ClassicAssert.AreEqual(
                convertWithSdk(value)!,
                parse(xml)!,
                $"{typeof(TSdk).Name}.{property.Name} ('{xml}')"
            );

            checkedValues++;
        }

        ClassicAssert.IsTrue(checkedValues > 0, $"no values found on {typeof(TSdk).Name}");

        foreach (TXl value in Enum.GetValues<TXl>())
        {
            ClassicAssert.AreEqual(
                ((IEnumValue)writeWithSdk(value)).Value,
                toXml(value),
                $"{typeof(TXl).Name}.{value}"
            );
        }
    }
}

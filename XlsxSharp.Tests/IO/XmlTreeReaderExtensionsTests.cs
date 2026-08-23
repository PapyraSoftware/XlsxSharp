using System.IO;
using System.Text;
using XlsxSharp.Excel.IO;
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal class XmlTreeReaderExtensionsTests
{
    private const string AttributeName = "test";

    [Test]
    public void GetDateTime_throws_when_attribute_is_not_present()
    {
        using XmlTreeReader reader = CreateReader("dummy");

        PartStructureException? ex = ClassicAssert.Throws<PartStructureException>(() =>
            reader.GetDateTime("nonexistent")
        );
        StringAssert.Contains(
            "XML doesn't contain a required attribute 'nonexistent'.",
            ex.Message
        );
    }

    [Test]
    public void GetXString_throws_when_attribute_is_not_present()
    {
        using XmlTreeReader reader = CreateReader("dummy");

        PartStructureException? ex = ClassicAssert.Throws<PartStructureException>(() =>
            reader.GetXString("nonexistent")
        );
        StringAssert.Contains(
            "XML doesn't contain a required attribute 'nonexistent'.",
            ex.Message
        );
    }

    [Test]
    [Arguments("&amp;", "&")]
    [Arguments("_x0009_", "\t")]
    [Arguments("_X0009_", "_X0009_")]
    [Arguments("Hello &lt;user&gt; - _x0045__x004F__x004C_", "Hello <user> - EOL")]
    public void GetOptionalXString_returns_XString_decoded_xml_decoded_text(
        string xmlText,
        string expectedValue
    )
    {
        using XmlTreeReader reader = CreateReader(xmlText);
        string? readValue = reader.GetOptionalXString(AttributeName);

        ClassicAssert.AreEqual(expectedValue, readValue);
    }

    [Test]
    [Arguments("00000000", 0u)]
    [Arguments("0G000000", null)]
    [Arguments(@"FFFFFFFF", 0xFFFFFFFF)]
    [Arguments(@"FFFFFFFF", 0xFFFFFFFF)]
    [Arguments("abcdef00", 0xABCDEF00)]
    [Arguments("0000000", null)]
    [Arguments(@"", null)]
    [Arguments(@"hello", null)]
    public void GetOptionalUIntHex_parses_8_hex_digits(string xmlText, uint? expectedValue)
    {
        using XmlTreeReader reader = CreateReader(xmlText);
        uint? readValue = reader.GetOptionalUIntHex(AttributeName);

        ClassicAssert.AreEqual(expectedValue, readValue);
    }

    private static XmlTreeReader CreateReader(string attributeValue)
    {
        string xmlContext = $"<element {AttributeName}=\"{attributeValue}\"/>";
        MemoryStream stream = new(Encoding.UTF8.GetBytes(xmlContext));
        XmlTreeReader reader = new(stream, new XmlToEnumMapper.Builder().Build(), false);
        reader.Open("element", string.Empty);
        return reader;
    }
}

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Excel.IO;
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

/// <summary>
/// Test that <see cref="XmlTreeReader"/> reads attributes per <see cref="XlsxSharp.Excel.LoadOptions.StrictAttributeParsing"/>
/// option. The invalid attribute values are either interpreted as a missing or throw an exception.
/// </summary>
[TestFixture]
[TestOf(typeof(XmlTreeReader))]
internal class XmlTreeReaderStrictParsingTests
{
    [Test]
    public void Reader_parses_bool_attributes_with_attribute_parsing_flag()
    {
        AssertStrictParsingFlag(
            """<element boolAttr="PRAVDA"/>""",
            reader => reader.GetOptionalBool("boolAttr")
        );
    }

    [Test]
    public void Reader_parses_dateTime_attributes_with_attribute_parsing_flag()
    {
        AssertStrictParsingFlag(
            """<element dateTimeAttr="tomorrow"/>""",
            reader => reader.GetOptionalDateTime("dateTimeAttr")
        );
    }

    [TestCase("pi")]
    [TestCase("1E+5000")]
    [TestCase("INF")]
    [TestCase("NaN")]
    public void Reader_parses_double_attributes_with_attribute_parsing_flag(string invalidValue)
    {
        AssertStrictParsingFlag(
            $"""<element doubleAttr="{invalidValue}"/>""",
            reader => reader.GetOptionalDouble("doubleAttr")
        );
    }

    [Test]
    public void Reader_parses_enum_attributes_with_attribute_parsing_flag()
    {
        AssertStrictParsingFlag(
            """<element enumAttr="triangle"/>""",
            reader => reader.GetOptionalEnum<XLBorderStyleValues>("enumAttr")
        );
    }

    [TestCase("zero")]
    [TestCase("5000000000000")]
    public void Reader_parses_int_attributes_with_attribute_parsing_flag(string invalidValue)
    {
        AssertStrictParsingFlag(
            $"""<element intAttr="{invalidValue}"/>""",
            reader => reader.GetOptionalInt("intAttr")
        );
    }

    [TestCase("zero")]
    [TestCase("-1")]
    [TestCase("4300000000")]
    [TestCase("10000000000000000000")] // Greater than long.MaxValue
    public void Reader_parses_uint_attributes_with_attribute_parsing_flag(string invalidValue)
    {
        AssertStrictParsingFlag(
            $"""<element uintAttr="{invalidValue}"/>""",
            reader => reader.GetOptionalUInt("uintAttr")
        );
    }

    private static void AssertStrictParsingFlag<T>(
        string xmlText,
        Func<XmlTreeReader, T> readAttribute
    )
    {
        AssertThrowsOnStrict(xmlText, readAttribute);
        AssertSkippedOnNonStrict(xmlText, readAttribute);
    }

    private static void AssertThrowsOnStrict<T>(
        string xmlText,
        Func<XmlTreeReader, T> readAttribute
    )
    {
        XAttribute attribute = XDocument.Parse(xmlText).Root!.Attributes().Single();
        string attributeName = attribute.Name.LocalName;
        string attributeValue = attribute.Value;

        using XmlTreeReader reader = new(
            new MemoryStream(Encoding.UTF8.GetBytes(xmlText)),
            XmlToEnumMapper.Instance,
            true
        );
        reader.Open("element", string.Empty);
        PartStructureException? ex = Assert.Throws<PartStructureException>(() =>
            readAttribute(reader)
        );
        Assert.That(
            ex?.Message,
            Does.StartWith(
                $"The attribute '{attributeName}' contains a value '{attributeValue}' that doesn't match expected format."
            )
        );
    }

    private static void AssertSkippedOnNonStrict<T>(
        string xmlText,
        Func<XmlTreeReader, T> readAttribute
    )
    {
        using XmlTreeReader reader = new(
            new MemoryStream(Encoding.UTF8.GetBytes(xmlText)),
            XmlToEnumMapper.Instance,
            false
        );
        reader.Open("element", string.Empty);
        T readAttributeValue = readAttribute(reader);
        Assert.IsNull(readAttributeValue);
    }
}

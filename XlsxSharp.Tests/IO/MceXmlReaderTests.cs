#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using XlsxSharp.IO;
using static XlsxSharp.IO.XmlTreeNodeType;

namespace XlsxSharp.Tests.IO;

internal class MceXmlReaderTests
{
    private const string MceNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string FooNs = "http://www.example.com/foo";
    private const string BarNs = "http://www.example.com/bar";

    [Test]
    [MethodDataSource<MceTestCase>(nameof(MceTestCase.ProcessingTestCases))]
    public void Processing_model_works_per_spec(MceTestCase test)
    {
        using MceXmlReader reader = CreateMceReader(test.Xml, test.AppConfig, test.Adee);

        foreach (IExpectedXmlNode expectedNode in test.ExpectedNodes)
        {
            ClassicAssert.IsTrue(reader.Read());
            expectedNode.AssertMatches(reader);
        }

        ClassicAssert.IsFalse(reader.Read());
        ClassicAssert.AreEqual(None, reader.NodeType);
        ClassicAssert.IsEmpty(reader.Value);
        ClassicAssert.IsEmpty(reader.LocalName);
        ClassicAssert.IsEmpty(reader.NamespaceUri);
    }

    [Test]
    public void Choice_requires_attribute_must_have_at_least_one_namespace()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}">
              <mc:Choice Requires="  "/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(reader, "MCE(2,4): Attribute Requires contains invalid value.");
    }

    [Test]
    public void Choice_requires_attribute_must_have_known_prefix()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}">
              <mc:Choice Requires="wrongPrefix"/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "MCE(2,4): Attribute Requires contains a namespace prefix wrongPrefix without a resolvable namespace."
        );
    }

    [Test]
    [Arguments(
        """
            <nonIgnorableNode/>
            <mc:Choice Requires="foo"/>
            """
    )]
    [Arguments(
        """
            <mc:Choice Requires="bar"/>
            <nonIgnorableNode/>
            <mc:Choice Requires="foo"/>
            """
    )]
    [Arguments(
        """
            <nonIgnorableNode/>
            <mc:Fallback/>
            """
    )]
    [Arguments(
        """
            <mc:Fallback/>
            <nonIgnorableNode/>
            """
    )]
    public void AlternateContent_doesnt_allow_unknown_elements_that_are_not_ignorable_as_children(
        string innerXml
    )
    {
        string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}">
              {innerXml}
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(xml, [FooNs]);

        AssertReadThrows(
            reader,
            "Element nonIgnorableNode must be from ignorable namespace, but isn't."
        );
    }

    [Test]
    public void AlternateContent_must_have_at_most_one_fallback()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}">
              <mc:Fallback/>
              <mc:Fallback/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(reader, "MCE(3,4): Found unexpected element Fallback.");
    }

    [Test]
    public void AlternateContent_fallback_must_be_after_all_choice_elements()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}">
              <mc:Fallback/>
              <mc:Choice Requires="foo"/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(reader, "MCE(3,4): Found unexpected element Choice.");
    }

    [Test]
    public void Ignorable_attribute_must_use_known_prefix()
    {
        const string xml = $"""<root xmlns:mc="{MceNs}" mc:Ignorable="unknownPrefix"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "MCE(1,2): Attribute Ignorable contains a namespace prefix unknownPrefix without a resolvable namespace."
        );
    }

    [Test]
    public void Ignorable_attribute_cant_use_mce_namespace()
    {
        const string xml = $"""<root xmlns:mc="{MceNs}" mc:Ignorable="mc"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "MCE(1,2): Attribute Ignorable contains namespace prefix for MCE, but that namespace is not allowed."
        );
    }

    [Test]
    public void ProcessContent_attribute_must_use_known_prefix()
    {
        const string xml = $"""<root xmlns:mc="{MceNs}" mc:ProcessContent="foo:*"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "MCE(1,2): Attribute ProcessContent contains a namespace prefix foo without a resolvable namespace."
        );
    }

    [Test]
    public void ProcessContent_attribute_cant_use_mce_namespace()
    {
        const string xml = $"""<root xmlns:mc="{MceNs}" mc:ProcessContent="mc:*"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "(1,2): Attribute ProcessContent contains namespace prefix for MCE, but that namespace is not allowed."
        );
    }

    [Test]
    public void ProcessContent_attribute_namespace_must_be_ignorable()
    {
        const string xml =
            $"""<root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:ProcessContent="foo:*"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(
            reader,
            "MCE(1,2): Attribute ProcessContent contains namespace http://www.example.com/foo that is not ignorable."
        );
    }

    [Test]
    [Arguments(":foo")]
    [Arguments("foo:")]
    [Arguments("foo:b:c")]
    [Arguments("foo:b+c")]
    public void ProcessContent_attribute_format_must_be_namespace_and_local_name(string value)
    {
        string xml =
            $"""<root xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:Ignorable="foo" mc:ProcessContent="{value}"/>""";
        using MceXmlReader reader = CreateMceReader(xml, []);

        AssertReadThrows(reader, "Attribute ProcessContent contains invalid value.");
    }

    [Test]
    public void MustUnderstand_namespaces_in_unwrapped_element_must_be_in_app_config()
    {
        const string xml = $"""
            <root xmlns:mc="{MceNs}" xmlns:bar="{BarNs}" mc:Ignorable="bar" mc:ProcessContent="bar:*">
              <bar:bar mc:MustUnderstand="bar"/>
            </root>
            """;
        using MceXmlReader reader = CreateMceReader(
            xml,
            [],
            mismatch: info => throw new InvalidOperationException($"Mismatch at {info.LineInfo}")
        );

        reader.Read();
        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            reader.Read()
        );
        ClassicAssert.AreEqual("Mismatch at 2:4", ex.Message);
    }

    [Test]
    public void MustUnderstand_namespaces_in_selected_choice_must_be_in_app_config()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}">
              <mc:Choice Requires="foo" mc:MustUnderstand="bar"/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(
            xml,
            [FooNs],
            mismatch: info => throw new InvalidOperationException($"Mismatch at {info.LineInfo}")
        );

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            reader.Read()
        );
        ClassicAssert.AreEqual("Mismatch at 2:4", ex.Message);
    }

    [Test]
    public void MustUnderstand_namespaces_in_alternate_content_must_be_in_app_config()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" mc:MustUnderstand="foo">
              <mc:Fallback/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(
            xml,
            [],
            mismatch: info => throw new InvalidOperationException($"Mismatch at {info.LineInfo}")
        );

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            reader.Read()
        );
        ClassicAssert.AreEqual("Mismatch at 1:2", ex.Message);
    }

    [Test]
    public void MustUnderstand_namespaces_in_selected_fallback_must_be_in_app_config()
    {
        const string xml = $"""
            <mc:AlternateContent xmlns:mc="{MceNs}" xmlns:foo="{FooNs}" xmlns:bar="{BarNs}">
              <mc:Choice Requires="foo" mc:MustUnderstand="foo"/>
              <mc:Fallback mc:MustUnderstand="bar"/>
            </mc:AlternateContent>
            """;
        using MceXmlReader reader = CreateMceReader(
            xml,
            [],
            mismatch: info => throw new InvalidOperationException($"Mismatch at {info.LineInfo}")
        );

        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            reader.Read()
        );
        ClassicAssert.AreEqual("Mismatch at 3:4", ex.Message);
    }

    private static MceXmlReader CreateMceReader(
        string xml,
        HashSet<string> appConfig,
        (string LocalName, string NamespaceUri)? adee = null,
        Action<MismatchInfo>? mismatch = null
    )
    {
        StringReader stringReader = new(xml);
        XmlReader xmlReader = XmlReader.Create(
            stringReader,
            new XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                DtdProcessing = DtdProcessing.Ignore,
            }
        );
        MceXmlReader mceReader = new(
            xmlReader,
            new MceSettings
            {
                ApplicationConfiguration = appConfig,
                AdeeLocalName = adee?.LocalName,
                AdeeNamespaceUri = adee?.NamespaceUri,
                SignalMismatch = mismatch,
            }
        );
        return mceReader;
    }

    private static void AssertReadThrows(MceXmlReader reader, string expectedMessage)
    {
        PartStructureException ex = ClassicAssert.Throws<PartStructureException>(() =>
            reader.Read()
        );
        StringAssert.StartsWith("MCE", ex.Message);
        ClassicAssert.IsTrue(ex.Message.EndsWith(expectedMessage, StringComparison.Ordinal));
    }
}

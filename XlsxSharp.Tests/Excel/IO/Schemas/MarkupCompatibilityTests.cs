using System.Xml.Linq;
using XlsxSharp.Excel.IO.Schemas;

namespace XlsxSharp.Tests.Excel.IO.Schemas;

public class MarkupCompatibilityTests
{
    private const string MceNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string ExtNs = "urn:test:extension";

    [Test]
    public void IgnorableAttributeAndItsNamespaceAreRemoved()
    {
        XElement root = XElement.Parse(
            $"""
            <root xmlns:mc="{MceNs}" xmlns:e="{ExtNs}" mc:Ignorable="e" e:foo="bar" plain="kept">
                <e:child/>
                <kept/>
            </root>
            """
        );

        MarkupCompatibility.Strip(root);

        ClassicAssert.IsNull(root.Attribute(XName.Get("Ignorable", MceNs)));
        ClassicAssert.IsNull(root.Attribute(XName.Get("foo", ExtNs)));
        ClassicAssert.AreEqual("kept", root.Attribute("plain")!.Value);
        ClassicAssert.IsNull(root.Element(XName.Get("child", ExtNs)));
        ClassicAssert.IsNotNull(root.Element("kept"));
    }

    [Test]
    public void IgnorableAppliesToTheWholeSubtreeNotJustDirectChildren()
    {
        XElement root = XElement.Parse(
            $"""
            <root xmlns:mc="{MceNs}" xmlns:e="{ExtNs}" mc:Ignorable="e">
                <wrapper>
                    <e:deep/>
                    <kept/>
                </wrapper>
            </root>
            """
        );

        MarkupCompatibility.Strip(root);

        XElement wrapper = root.Element("wrapper")!;
        ClassicAssert.IsNull(wrapper.Element(XName.Get("deep", ExtNs)));
        ClassicAssert.IsNotNull(wrapper.Element("kept"));
    }

    [Test]
    public void AlternateContentIsReplacedByItsFallback()
    {
        XElement root = XElement.Parse(
            $"""
            <root xmlns:mc="{MceNs}" xmlns:e="{ExtNs}">
                <mc:AlternateContent>
                    <mc:Choice Requires="e"><e:new/></mc:Choice>
                    <mc:Fallback><old/></mc:Fallback>
                </mc:AlternateContent>
            </root>
            """
        );

        MarkupCompatibility.Strip(root);

        ClassicAssert.IsNull(root.Element(XName.Get("AlternateContent", MceNs)));
        ClassicAssert.IsNotNull(root.Element("old"));
        ClassicAssert.IsNull(root.Element(XName.Get("new", ExtNs)));
    }

    [Test]
    public void AlternateContentWithoutAFallbackIsDroppedEntirely()
    {
        XElement root = XElement.Parse(
            $"""
            <root xmlns:mc="{MceNs}" xmlns:e="{ExtNs}">
                <mc:AlternateContent>
                    <mc:Choice Requires="e"><e:new/></mc:Choice>
                </mc:AlternateContent>
                <kept/>
            </root>
            """
        );

        MarkupCompatibility.Strip(root);

        ClassicAssert.IsFalse(root.Elements().Any(e => e.Name.LocalName != "kept"));
    }

    [Test]
    public void AnExtWithAUriThatBecomesEmptyIsRemoved()
    {
        // Mirrors what a real workbook.xml carries: an extLst/ext whose one child is in a
        // namespace the root marks ignorable - once that child is stripped, the wrapper itself
        // has to go too, or schema validation would reject it for having no content at all.
        XElement root = XElement.Parse(
            $"""
            <root xmlns:mc="{MceNs}" xmlns:e="{ExtNs}" mc:Ignorable="e">
                <extLst>
                    <ext uri="11111111-1111-1111-1111-111111111111"><e:feature/></ext>
                </extLst>
            </root>
            """
        );

        MarkupCompatibility.Strip(root);

        ClassicAssert.IsFalse(root.Element("extLst")!.Elements().Any());
    }

    [Test]
    public void AnUnrelatedElementNamedExtWithNoUriIsNeverRemoved()
    {
        // DrawingML's anchor size element is also called "ext" and is legitimately childless by
        // its own type (just cx/cy attributes) - it must survive untouched.
        XElement root = XElement.Parse("<root><ext cx=\"100\" cy=\"200\"/></root>");

        MarkupCompatibility.Strip(root);

        ClassicAssert.IsNotNull(root.Element("ext"));
    }
}

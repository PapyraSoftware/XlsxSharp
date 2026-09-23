using System.Text;
using XlsxSharp.Excel.IO;

namespace XlsxSharp.Tests.Excel.IO;

public class SharedStringTableReaderTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    public void PlainEntriesAreKeptAsTheirDecodedText()
    {
        SharedString[] strings = Read(
            """
            <si><t>abc</t></si>
            <si><t>line_x000D_break</t></si>
            <si><t xml:space="preserve"> padded </t></si>
            <si><t> </t></si>
            <si/>
            """
        );

        ClassicAssert.AreEqual(5, strings.Length);
        ClassicAssert.IsTrue(strings.All(s => s.Item is null));
        ClassicAssert.AreEqual("abc", strings[0].Text);
        ClassicAssert.AreEqual("line\rbreak", strings[1].Text);
        ClassicAssert.AreEqual(" padded ", strings[2].Text);

        // Whitespace outside xml:space="preserve" is insignificant.
        ClassicAssert.AreEqual(string.Empty, strings[3].Text);
        ClassicAssert.AreEqual(string.Empty, strings[4].Text);
    }

    [Test]
    public void RunsKeepTheirTextAndFormatting()
    {
        SharedString entry = Read(
                """
                <si>
                  <r><t>plain </t></r>
                  <r><rPr><b/><sz val="12"/></rPr><t xml:space="preserve">bold </t></r>
                </si>
                """
            )
            .Single();

        StringItem item = entry.Item!;
        ClassicAssert.AreEqual(2, item.Runs!.Count);
        ClassicAssert.IsNull(item.Runs[0].Properties);
        ClassicAssert.AreEqual("plain ", item.Runs[0].Text);
        ClassicAssert.AreEqual("bold ", item.Runs[1].Text);
        ClassicAssert.AreEqual("rPr", item.Runs[1].Properties!.Name.LocalName);
        ClassicAssert.AreEqual(2, item.Runs[1].Properties!.Elements().Count());
    }

    [Test]
    public void APhoneticGuideMakesAnEntryMoreThanPlainText()
    {
        SharedString entry = Read(
                """
                <si>
                  <t>東京</t>
                  <rPh sb="0" eb="2"><t>トウキョウ</t></rPh>
                  <phoneticPr fontId="1" type="Hiragana"/>
                </si>
                """
            )
            .Single();

        StringItem item = entry.Item!;
        ClassicAssert.AreEqual("東京", item.Text);
        ClassicAssert.AreEqual(("トウキョウ", 0, 2), item.Phonetics!.Single());
        ClassicAssert.AreEqual("Hiragana", item.PhoneticProperties!.Attribute("type")!.Value);
    }

    [Test]
    public void ContentOutsideTheMainNamespaceIsSkipped()
    {
        SharedString[] strings = Read(
            """
            <si xmlns:w14="urn:example:w14"><w14:extra><t>not this</t></w14:extra><t>this</t></si>
            <si><t>next</t></si>
            """
        );

        ClassicAssert.AreEqual("this|next", string.Join("|", strings.Select(s => s.Text)));
    }

    [Test]
    public void AnEmptyTableHasNoEntries() => ClassicAssert.AreEqual(0, Read(string.Empty).Length);

    private static SharedString[] Read(string entries)
    {
        string xml = $"""<sst xmlns="{Main}">{entries}</sst>""";
        using MemoryStream stream = new(Encoding.UTF8.GetBytes(xml));
        return SharedStringTableReader.Read(stream);
    }
}

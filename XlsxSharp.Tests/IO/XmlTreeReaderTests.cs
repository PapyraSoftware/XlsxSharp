using System.Text;
using XlsxSharp.Excel.IO;
using XlsxSharp.IO;

namespace XlsxSharp.Tests.IO;

internal class XmlTreeReaderTests
{
    [Test]
    public void Can_transparently_processes_MCE()
    {
        const string xml = $"""
            <font xmlns="{OpenXmlConst.Main2006SsNs}"
                   xmlns:mc="{OpenXmlConst.MarkupCompatibilityNs}">
              <mc:AlternateContent>
                <mc:Choice xmlns:cs="http://example.com/custom" Requires="cs">
                  <cs:bold weight="10"/>
                </mc:Choice>
                <mc:Fallback>
                  <b/>
                </mc:Fallback>
              </mc:AlternateContent>
            </font>
            """;
        using XmlTreeReader reader = new(
            new MemoryStream(Encoding.UTF8.GetBytes(xml)),
            XmlToEnumMapper.Instance,
            true
        );
        reader.Open("font", OpenXmlConst.Main2006SsNs);
        reader.Open("b", OpenXmlConst.Main2006SsNs);
        reader.Close("b", OpenXmlConst.Main2006SsNs);
        reader.Close("font", OpenXmlConst.Main2006SsNs);
    }
}

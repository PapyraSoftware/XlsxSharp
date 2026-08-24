using XlsxSharp.Excel;
using XlsxSharp.Utils;

namespace XlsxSharp.Tests.Excel.Misc;

public class XmlEncoderTest
{
    [Test]
    public void TestControlChars()
    {
        ClassicAssert.AreEqual(
            "_x0001_ _x0002_ _x0003_ _x0004_",
            XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004")
        );
        ClassicAssert.AreEqual(
            "_x0005_ _x0006_ _x0007_ _x0008_",
            XmlEncoder.EncodeString("\u0005 \u0006 \u0007 \u0008")
        );
    }

    [Test]
    public void AstralUnicodeCharsAreWrittenWithoutOpenXmlEncoding()
    {
        using StreamReader sr = new(
            TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Other\Unicode\let_it_go_in_emoji.txt")
            )
        );
        string surrogateEmoji = sr.ReadToEnd();

        TestHelper.CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();

                IXLCell cell = ws.FirstCell();
                cell.Value = "This emoji version of Let It Go from Frozen:";
                cell.CellBelow().Value = surrogateEmoji;
            },
            @"Other\Unicode\let_it_go_in_emoji-outputfile.xlsx"
        );
    }
}

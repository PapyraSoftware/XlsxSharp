using DocumentFormat.OpenXml.Packaging;

namespace XlsxSharp.Tests;

public class OpenXmlTests
{
    [Test]
    [Skip("Workaround has been included in XlsxSharp")]
    public void SetPackagePropertiesEntryToNullWithOpenXml()
    {
        // Fixed in .NET Standard 2.1
        // See:
        //      https://github.com/OfficeDev/Open-XML-SDK/issues/235
        //      https://github.com/dotnet/corefx/issues/23795
        using (
            Stream stream = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")
            )
        )
        using (MemoryStream ms = new())
        {
            stream.CopyTo(ms);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, true))
            {
                document.PackageProperties.Creator = null;
            }
        }
    }
}

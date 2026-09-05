using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

/// <summary>
/// Checks the packaging layer against real workbooks rather than against packages it built
/// itself: it has to see what <c>DocumentFormat.OpenXml.Packaging</c> sees, and what it writes
/// has to stay readable by both the SDK and XlsxSharp. These tests are what says the layer can
/// take over from the SDK.
/// </summary>
public class OpcPackageInteropTests
{
    [Test]
    public void SeesTheSamePartsAndContentTypesAsTheSdk()
    {
        using MemoryStream stream = CreateWorkbook();

        stream.Position = 0;
        Dictionary<string, string> expected;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
        {
            expected = document
                .GetAllParts()
                .ToDictionary(p => p.Uri.OriginalString, p => p.ContentType, OpcPartName.Comparer);
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);
        Dictionary<string, string> actual = package.Parts.ToDictionary(
            p => p.Name,
            p => p.ContentType,
            OpcPartName.Comparer
        );

        ClassicAssert.AreEqual(
            expected.OrderBy(x => x.Key, OpcPartName.Comparer).ToList(),
            actual.OrderBy(x => x.Key, OpcPartName.Comparer).ToList()
        );
    }

    [Test]
    public void ResolvesTheSameRelationshipTargetsAsTheSdk()
    {
        using MemoryStream stream = CreateWorkbook();

        stream.Position = 0;
        List<(string Source, string Id, string Target)> expected = [];
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
        {
            foreach (
                IdPartPair pair in document
                    .Parts.Concat(document.WorkbookPart!.Parts)
                    .OrderBy(p => p.RelationshipId, StringComparer.Ordinal)
            )
            {
                expected.Add(
                    (Source(pair), pair.RelationshipId, pair.OpenXmlPart.Uri.OriginalString)
                );
            }
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);

        OpcPart workbookPart = package.GetPart("/xl/workbook.xml");
        List<(string Source, string Id, string Target)> actual = [];
        foreach (
            OpcRelationship relationship in package.Relationships.OrderBy(
                r => r.Id,
                StringComparer.Ordinal
            )
        )
        {
            actual.Add((string.Empty, relationship.Id, relationship.TargetPartName!));
        }

        foreach (
            OpcRelationship relationship in workbookPart.Relationships.OrderBy(
                r => r.Id,
                StringComparer.Ordinal
            )
        )
        {
            actual.Add(("/xl/workbook.xml", relationship.Id, relationship.TargetPartName!));
        }

        ClassicAssert.AreEqual(
            expected.OrderBy(x => x.Source + x.Id, StringComparer.Ordinal).ToList(),
            actual.OrderBy(x => x.Source + x.Id, StringComparer.Ordinal).ToList()
        );

        static string Source(IdPartPair pair) =>
            pair.OpenXmlPart.Uri.OriginalString.StartsWith("/xl/", StringComparison.Ordinal)
            && pair.OpenXmlPart.Uri.OriginalString != "/xl/workbook.xml"
                ? "/xl/workbook.xml"
                : string.Empty;
    }

    [Test]
    public void APackageRewrittenByTheLayerStaysReadableByTheSdk()
    {
        using MemoryStream original = CreateWorkbook();

        using MemoryStream rewritten = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, writable: true))
        {
            package.SaveTo(rewritten);
        }

        rewritten.Position = 0;
        using SpreadsheetDocument document = SpreadsheetDocument.Open(rewritten, false);

        ClassicAssert.IsNotNull(document.WorkbookPart);
        ClassicAssert.IsNotNull(document.WorkbookPart!.Workbook);
        ClassicAssert.AreEqual(1, document.WorkbookPart.WorksheetParts.Count());
    }

    [Test]
    public void APackageRewrittenByTheLayerStaysReadableByXlsxSharp()
    {
        using MemoryStream original = CreateWorkbook();

        using MemoryStream rewritten = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, writable: true))
        {
            package.SaveTo(rewritten);
        }

        rewritten.Position = 0;
        using XLWorkbook workbook = new(rewritten);

        IXLWorksheet worksheet = workbook.Worksheet("Data");
        ClassicAssert.AreEqual("Hello", worksheet.Cell("A1").GetString());
        ClassicAssert.AreEqual(42, worksheet.Cell("B1").GetDouble());
    }

    /// <summary>
    /// A workbook with enough in it to exercise several parts and both package level and part
    /// level relationships.
    /// </summary>
    private static MemoryStream CreateWorkbook()
    {
        MemoryStream stream = new();
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet worksheet = workbook.AddWorksheet("Data");
            worksheet.Cell("A1").Value = "Hello";
            worksheet.Cell("B1").Value = 42;
            worksheet.Cell("A2").Style.Fill.BackgroundColor = XLColor.Red;
            workbook.Properties.Author = "XlsxSharp";
            workbook.SaveAs(stream);
        }

        return stream;
    }
}

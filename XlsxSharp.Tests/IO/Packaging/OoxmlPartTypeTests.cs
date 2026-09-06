using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

/// <summary>
/// The part type table is a transcription of what the SDK's typed part classes know. These tests
/// check it against the SDK rather than against itself, so that a wrong content type or
/// relationship type shows up here and not as a workbook Excel refuses to open.
/// </summary>
public class OoxmlPartTypeTests
{
    [Test]
    public void EveryPartOfARichWorkbookIsCoveredByTheTable()
    {
        using MemoryStream stream = RichWorkbook();

        stream.Position = 0;
        Dictionary<string, string> sdkContentTypes;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
        {
            sdkContentTypes = document
                .GetAllParts()
                .ToDictionary(p => p.Uri.OriginalString, p => p.ContentType, OpcPartName.Comparer);
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);

        // Walk the package the way the table says to, and check that each part found that way has
        // the content type the SDK reports for it.
        OpcPart workbook = package.PartOfType(OoxmlPartTypes.Workbook)!;
        ClassicAssert.IsNotNull(workbook);
        AssertMatchesSdk(workbook, OoxmlPartTypes.Workbook);

        foreach (
            (OoxmlPartType partType, int expectedCount) in new[]
            {
                (OoxmlPartTypes.Worksheet, 2),
                (OoxmlPartTypes.Styles, 1),
                (OoxmlPartTypes.Theme, 1),
                (OoxmlPartTypes.SharedStringTable, 1),
            }
        )
        {
            List<OpcPart> parts = [.. workbook.PartsOfType(partType)];
            ClassicAssert.AreEqual(
                expectedCount,
                parts.Count,
                $"unexpected number of {partType.PathTemplate} parts"
            );

            parts.ForEach(p => AssertMatchesSdk(p, partType));
        }

        OpcPart sheet = workbook.PartsOfType(OoxmlPartTypes.Worksheet).First();
        AssertMatchesSdk(sheet.PartsOfType(OoxmlPartTypes.Table).Single(), OoxmlPartTypes.Table);
        AssertMatchesSdk(
            sheet.PartsOfType(OoxmlPartTypes.Drawing).Single(),
            OoxmlPartTypes.Drawing
        );

        AssertMatchesSdk(
            package.PartOfType(OoxmlPartTypes.ExtendedFileProperties)!,
            OoxmlPartTypes.ExtendedFileProperties
        );

        void AssertMatchesSdk(OpcPart part, OoxmlPartType partType)
        {
            ClassicAssert.IsTrue(
                sdkContentTypes.ContainsKey(part.Name),
                $"the SDK does not know a part named {part.Name}"
            );

            ClassicAssert.AreEqual(
                sdkContentTypes[part.Name],
                part.ContentType,
                $"content type of {part.Name}"
            );

            ClassicAssert.AreEqual(
                partType.ContentType,
                part.ContentType,
                $"the table's content type for {part.Name}"
            );
        }
    }

    [Test]
    public void PartsAreFoundWhereTheSdkPutsThem()
    {
        using MemoryStream stream = RichWorkbook();

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);
        OpcPart workbook = package.PartOfType(OoxmlPartTypes.Workbook)!;

        // The templates say where a new part goes; a workbook written by XlsxSharp has to already
        // agree with them, otherwise adding a part would collide with an existing one.
        ClassicAssert.AreEqual("/xl/workbook.xml", workbook.Name);
        ClassicAssert.AreEqual("/xl/styles.xml", workbook.PartOfType(OoxmlPartTypes.Styles)!.Name);

        ClassicAssert.AreEqual(
            "/xl/sharedStrings.xml",
            workbook.PartOfType(OoxmlPartTypes.SharedStringTable)!.Name
        );

        ClassicAssert.IsTrue(
            workbook
                .PartsOfType(OoxmlPartTypes.Worksheet)
                .All(p => p.Name.StartsWith("/xl/worksheets/sheet", StringComparison.Ordinal))
        );
    }

    [Test]
    public void AddingNumberedPartsPicksTheNextFreeName()
    {
        using OpcPackage package = OpcPackage.Create();
        (OpcPart workbook, _) = package.AddPartOfType(OoxmlPartTypes.Workbook);

        (OpcPart first, string firstId) = workbook.AddPartOfType(package, OoxmlPartTypes.Worksheet);

        (OpcPart second, string secondId) = workbook.AddPartOfType(
            package,
            OoxmlPartTypes.Worksheet
        );

        ClassicAssert.AreEqual("/xl/worksheets/sheet1.xml", first.Name);
        ClassicAssert.AreEqual("/xl/worksheets/sheet2.xml", second.Name);
        ClassicAssert.AreNotEqual(firstId, secondId);

        // The relationship goes on the workbook part, not on the package.
        ClassicAssert.AreEqual(2, workbook.PartsOfType(OoxmlPartTypes.Worksheet).Count());
        ClassicAssert.AreEqual(0, package.PartsOfType(OoxmlPartTypes.Worksheet).Count());
        ClassicAssert.AreEqual(first.Name, workbook.GetRelatedPart(firstId).Name);
    }

    [Test]
    public void AnUnnumberedPartTypeAlwaysLandsOnTheSameName()
    {
        using OpcPackage package = OpcPackage.Create();
        (OpcPart workbook, _) = package.AddPartOfType(OoxmlPartTypes.Workbook);
        (OpcPart styles, _) = workbook.AddPartOfType(package, OoxmlPartTypes.Styles);

        ClassicAssert.AreEqual("/xl/styles.xml", styles.Name);
        ClassicAssert.IsFalse(OoxmlPartTypes.Styles.IsNumbered);
        ClassicAssert.IsTrue(OoxmlPartTypes.Worksheet.IsNumbered);
    }

    [Test]
    public void ImagesCarryTheirOwnContentType()
    {
        using OpcPackage package = OpcPackage.Create();
        (OpcPart workbook, _) = package.AddPartOfType(OoxmlPartTypes.Workbook);
        (OpcPart image, _) = workbook.AddPartOfType(
            package,
            OoxmlPartTypes.Image,
            contentType: "image/jpeg",
            partName: "/xl/media/image1.jpeg"
        );

        ClassicAssert.AreEqual("image/jpeg", image.ContentType);
    }

    /// <summary>A workbook with enough in it to bring most of the part kinds into the package.</summary>
    private static MemoryStream RichWorkbook()
    {
        MemoryStream stream = new();
        using (XLWorkbook workbook = new())
        {
            IXLWorksheet first = workbook.AddWorksheet("First");
            first.Cell("A1").Value = "Name";
            first.Cell("A2").Value = "Ada";
            first.Range("A1:A2").CreateTable();

            using (
                Stream image = System
                    .Reflection.Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("XlsxSharp.Tests.Resource.Images.ImageHandling.png")!
            )
            {
                first.AddPicture(image).MoveTo(first.Cell("C3"));
            }

            IXLWorksheet second = workbook.AddWorksheet("Second");
            second.Cell("A1").Value = "Shared";

            workbook.SaveAs(stream);
        }

        return stream;
    }
}

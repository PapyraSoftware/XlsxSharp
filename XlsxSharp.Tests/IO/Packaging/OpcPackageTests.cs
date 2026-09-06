using System.Text;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

public class OpcPackageTests
{
    private const string WorkbookContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    private const string WorksheetContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    private const string OfficeDocumentRel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

    private const string WorksheetRel =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    [Test]
    public void PartsRoundTripThroughTheZip()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
            Write(workbook, "<workbook/>");
            package.Relationships.Add(workbook.Name, OfficeDocumentRel);
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        OpcPart part = reopened.GetPart("/xl/workbook.xml");
        ClassicAssert.AreEqual(WorkbookContentType, part.ContentType);
        ClassicAssert.AreEqual("<workbook/>", Read(part));
    }

    [Test]
    public void RelationshipsRoundTripAndResolveToParts()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
            OpcPart sheet = package.AddPart("/xl/worksheets/sheet1.xml", WorksheetContentType);
            Write(workbook, "<workbook/>");
            Write(sheet, "<worksheet/>");

            package.Relationships.Add(workbook.Name, OfficeDocumentRel);
            workbook.Relationships.Add(sheet.Name, WorksheetRel, "rId7");
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        OpcPart workbookPart = reopened.GetRelatedParts(OfficeDocumentRel).Single();
        ClassicAssert.AreEqual("/xl/workbook.xml", workbookPart.Name);

        OpcRelationship sheetRelationship = workbookPart.Relationships.GetById("rId7");

        // The target is written relative to the folder of the part declaring it.
        ClassicAssert.AreEqual("worksheets/sheet1.xml", sheetRelationship.Target);
        ClassicAssert.AreEqual("/xl/worksheets/sheet1.xml", sheetRelationship.TargetPartName);
        ClassicAssert.AreEqual("<worksheet/>", Read(workbookPart.GetRelatedPart("rId7")));
        ClassicAssert.AreSame(
            workbookPart.GetRelatedPart("rId7"),
            workbookPart.GetRelatedPartOrDefault("rId7")
        );
        ClassicAssert.IsNull(workbookPart.GetRelatedPartOrDefault("rIdMissing"));
    }

    [Test]
    public void GetRelatedPartOrDefaultIsNullForAnExternalRelationship()
    {
        using OpcPackage package = OpcPackage.Create();
        OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
        workbook.Relationships.AddExternal(
            "https://example.com",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            "rId1"
        );

        ClassicAssert.IsNull(workbook.GetRelatedPartOrDefault("rId1"));
    }

    [Test]
    public void GeneratedRelationshipIdsSkipTheOnesAlreadyTaken()
    {
        using OpcPackage package = OpcPackage.Create();
        OpcPart a = package.AddPart("/a.xml", OpcContentType.Xml);
        OpcPart b = package.AddPart("/b.xml", OpcContentType.Xml);

        package.Relationships.Add(a.Name, WorksheetRel, "rId1");
        OpcRelationship generated = package.Relationships.Add(b.Name, WorksheetRel);

        ClassicAssert.AreEqual("rId2", generated.Id);
    }

    [Test]
    public void AddingADuplicateRelationshipIdThrows()
    {
        using OpcPackage package = OpcPackage.Create();
        OpcPart a = package.AddPart("/a.xml", OpcContentType.Xml);

        package.Relationships.Add(a.Name, WorksheetRel, "rId1");
        ClassicAssert.Throws<OpcException>(() =>
            package.Relationships.Add(a.Name, WorksheetRel, "rId1")
        );
    }

    [Test]
    public void ExternalRelationshipsKeepTheirTargetAndAreNotParts()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart sheet = package.AddPart("/xl/worksheets/sheet1.xml", WorksheetContentType);
            Write(sheet, "<worksheet/>");
            sheet.Relationships.AddExternal("https://example.com/", "http://example.com/hyperlink");
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        OpcRelationship relationship = reopened
            .GetPart("/xl/worksheets/sheet1.xml")
            .Relationships.Single();

        ClassicAssert.AreEqual(OpcTargetMode.External, relationship.TargetMode);
        ClassicAssert.AreEqual("https://example.com/", relationship.Target);
        ClassicAssert.IsNull(relationship.TargetPartName);
    }

    [Test]
    public void DeletingAPartAlsoDropsTheRelationshipsPointingAtIt()
    {
        using OpcPackage package = OpcPackage.Create();
        OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
        OpcPart sheet = package.AddPart("/xl/worksheets/sheet1.xml", WorksheetContentType);
        workbook.Relationships.Add(sheet.Name, WorksheetRel);

        package.DeletePart(sheet.Name);

        ClassicAssert.IsFalse(package.TryGetPart(sheet.Name, out _));
        ClassicAssert.AreEqual(0, workbook.Relationships.Count);
    }

    [Test]
    public void ContentTypesUseTheDefaultWhenItMatchesAndAnOverrideOtherwise()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            // Same type as the "xml" default that every package starts with.
            OpcPart plain = package.AddPart("/plain.xml", OpcContentType.Xml);

            // A different type for the same extension, so this one needs an override.
            OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
            Write(plain, "<plain/>");
            Write(workbook, "<workbook/>");
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        ClassicAssert.AreEqual(OpcContentType.Xml, reopened.GetPart("/plain.xml").ContentType);
        ClassicAssert.AreEqual(
            WorkbookContentType,
            reopened.GetPart("/xl/workbook.xml").ContentType
        );
    }

    [Test]
    public void BinaryPartsGetADefaultForTheirExtension()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart image = package.AddPart("/xl/media/image1.png", "image/png");
            image.GetWriteStream().Dispose();
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);
        ClassicAssert.AreEqual("image/png", reopened.GetPart("/xl/media/image1.png").ContentType);
    }

    [Test]
    public void RewritingAPartReplacesItsContent()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart part = package.AddPart("/a.xml", OpcContentType.Xml);
            Write(part, "<first/>");
            Write(part, "<second/>");
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);
        ClassicAssert.AreEqual("<second/>", Read(reopened.GetPart("/a.xml")));
    }

    [Test]
    public void LengthIsTheContentSizeForAPartWrittenInMemory()
    {
        using OpcPackage package = OpcPackage.Create();
        OpcPart part = package.AddPart("/a.xml", OpcContentType.Xml);
        ClassicAssert.AreEqual(0, part.Length);

        Write(part, "<first/>");
        ClassicAssert.AreEqual(Encoding.UTF8.GetByteCount("<first/>"), part.Length);
    }

    [Test]
    public void LengthIsTheUncompressedSizeForAPartBackedByTheZipWithoutReadingIt()
    {
        const string content = "<first/>";

        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            Write(package.AddPart("/a.xml", OpcContentType.Xml), content);
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);
        OpcPart part = reopened.GetPart("/a.xml");

        // Asserted before GetReadStream() is ever called on this part: a part read back from the
        // ZIP reports its length straight from the entry's own directory record, not by
        // decompressing it first.
        ClassicAssert.AreEqual(Encoding.UTF8.GetByteCount(content), part.Length);
        ClassicAssert.AreEqual(content, Read(part));
    }

    [Test]
    public void AnUntouchedPartSurvivesAReadModifyWriteCycle()
    {
        using MemoryStream original = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            Write(package.AddPart("/keep.xml", OpcContentType.Xml), "<keep/>");
            Write(package.AddPart("/change.xml", OpcContentType.Xml), "<before/>");
            package.SaveTo(original);
        }

        using MemoryStream rewritten = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, writable: true))
        {
            Write(package.GetPart("/change.xml"), "<after/>");
            package.SaveTo(rewritten);
        }

        rewritten.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(rewritten);
        ClassicAssert.AreEqual("<keep/>", Read(reopened.GetPart("/keep.xml")));
        ClassicAssert.AreEqual("<after/>", Read(reopened.GetPart("/change.xml")));
    }

    [Test]
    public void AReadOnlyPackageRefusesModification()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            Write(package.AddPart("/a.xml", OpcContentType.Xml), "<a/>");
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        ClassicAssert.IsTrue(reopened.IsReadOnly);
        ClassicAssert.Throws<InvalidOperationException>(() =>
            reopened.AddPart("/b.xml", OpcContentType.Xml)
        );
        ClassicAssert.Throws<InvalidOperationException>(() =>
            reopened.GetPart("/a.xml").GetWriteStream()
        );
    }

    [Test]
    public void OpeningSomethingThatIsNotAZipThrows()
    {
        using MemoryStream stream = new(Encoding.UTF8.GetBytes("not a zip"));
        ClassicAssert.Throws<OpcException>(() => OpcPackage.Open(stream));
    }

    [Test]
    public void OpeningAZipWithoutContentTypesThrows()
    {
        using MemoryStream stream = new();
        using (
            System.IO.Compression.ZipArchive archive = new(
                stream,
                System.IO.Compression.ZipArchiveMode.Create,
                leaveOpen: true
            )
        )
        {
            archive.CreateEntry("xl/workbook.xml");
        }

        stream.Position = 0;
        ClassicAssert.Throws<OpcException>(() => OpcPackage.Open(stream));
    }

    [Test]
    public void PackagesRoundTripThroughAFile()
    {
        string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");
        try
        {
            using (OpcPackage package = OpcPackage.Create(path))
            {
                OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
                Write(workbook, "<workbook/>");
                package.Relationships.Add(workbook.Name, OfficeDocumentRel);
            }

            using (OpcPackage package = OpcPackage.Open(path, writable: true))
            {
                Write(package.GetPart("/xl/workbook.xml"), "<edited/>");
            }

            using OpcPackage reopened = OpcPackage.Open(path);
            ClassicAssert.AreEqual("<edited/>", Read(reopened.GetPart("/xl/workbook.xml")));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Test]
    public void PackagesRoundTripThroughAWritableStream()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart workbook = package.AddPart("/xl/workbook.xml", WorkbookContentType);
            Write(workbook, "<workbook/>");
            package.Relationships.Add(workbook.Name, OfficeDocumentRel);
            package.SaveTo(stream);
        }

        // The same stream the package is opened from is also where a writable open saves back to
        // - it must not still be readable by anyone still holding the entries it was opened with.
        using (OpcPackage package = OpcPackage.Open(stream, writable: true))
        {
            Write(package.GetPart("/xl/workbook.xml"), "<edited/>");
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);
        ClassicAssert.AreEqual("<edited/>", Read(reopened.GetPart("/xl/workbook.xml")));
    }

    private static void Write(OpcPart part, string content)
    {
        using Stream stream = part.GetWriteStream();
        stream.Write(Encoding.UTF8.GetBytes(content));
    }

    private static string Read(OpcPart part)
    {
        using Stream stream = part.GetReadStream();
        using StreamReader reader = new(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}

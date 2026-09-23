using System.IO.Packaging;
using XlsxSharp.Excel;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

public class OpcPackagePropertiesTests
{
    [Test]
    public void PropertiesRoundTripThroughTheCorePart()
    {
        DateTime created = new(2024, 3, 17, 8, 30, 0, DateTimeKind.Utc);
        DateTime modified = new(2025, 11, 2, 19, 45, 12, DateTimeKind.Utc);

        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            package.Properties.Creator = "Ada";
            package.Properties.LastModifiedBy = "Grace";
            package.Properties.Title = "Quarterly";
            package.Properties.Subject = "Numbers";
            package.Properties.Description = "A description";
            package.Properties.Keywords = "one two";
            package.Properties.Category = "Reports";
            package.Properties.ContentStatus = "Draft";
            package.Properties.Created = created;
            package.Properties.Modified = modified;
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        ClassicAssert.AreEqual("Ada", reopened.Properties.Creator);
        ClassicAssert.AreEqual("Grace", reopened.Properties.LastModifiedBy);
        ClassicAssert.AreEqual("Quarterly", reopened.Properties.Title);
        ClassicAssert.AreEqual("Numbers", reopened.Properties.Subject);
        ClassicAssert.AreEqual("A description", reopened.Properties.Description);
        ClassicAssert.AreEqual("one two", reopened.Properties.Keywords);
        ClassicAssert.AreEqual("Reports", reopened.Properties.Category);
        ClassicAssert.AreEqual("Draft", reopened.Properties.ContentStatus);
        ClassicAssert.AreEqual(created, reopened.Properties.Created);
        ClassicAssert.AreEqual(modified, reopened.Properties.Modified);
    }

    [Test]
    public void APropertyThatWasNeverSetStaysNull()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            package.Properties.Creator = "Ada";
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);

        ClassicAssert.AreEqual("Ada", reopened.Properties.Creator);
        ClassicAssert.IsNull(reopened.Properties.Title);
        ClassicAssert.IsNull(reopened.Properties.Created);
    }

    [Test]
    public void TheCorePartIsOnlyWrittenWhenAPropertyWasSet()
    {
        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            package.AddPart("/a.xml", OpcContentType.Xml).GetWriteStream().Dispose();

            // Reading must not create the part.
            _ = package.Properties.Creator;
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(stream);
        ClassicAssert.IsFalse(reopened.TryGetPart("/docProps/core.xml", out _));
    }

    [Test]
    public void AnotherImplementationReadsBackWhatTheLayerWrote()
    {
        DateTime created = new(2024, 3, 17, 8, 30, 0, DateTimeKind.Utc);

        using MemoryStream stream = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart workbook = MinimalWorkbook(package);
            package.Relationships.Add(
                workbook.Name,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
            );

            package.Properties.Creator = "Ada";
            package.Properties.Title = "Quarterly";
            package.Properties.Created = created;
            package.SaveTo(stream);
        }

        stream.Position = 0;
        using Package independent = Package.Open(stream, FileMode.Open, FileAccess.Read);

        ClassicAssert.AreEqual("Ada", independent.PackageProperties.Creator);
        ClassicAssert.AreEqual("Quarterly", independent.PackageProperties.Title);
        ClassicAssert.AreEqual(created, independent.PackageProperties.Created?.ToUniversalTime());
    }

    [Test]
    public void TheLayerReadsWhatAnotherImplementationWrote()
    {
        DateTime created = new(2024, 3, 17, 8, 30, 0, DateTimeKind.Utc);

        using MemoryStream stream = new();
        using (Package independent = Package.Open(stream, FileMode.Create, FileAccess.ReadWrite))
        {
            independent.PackageProperties.Creator = "Ada";
            independent.PackageProperties.Title = "Quarterly";
            independent.PackageProperties.Category = "Reports";
            independent.PackageProperties.Created = created;
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);

        ClassicAssert.AreEqual("Ada", package.Properties.Creator);
        ClassicAssert.AreEqual("Quarterly", package.Properties.Title);
        ClassicAssert.AreEqual("Reports", package.Properties.Category);
        ClassicAssert.AreEqual(created, package.Properties.Created);
    }

    [Test]
    public void TheLayerReadsWhatXlsxSharpWrote()
    {
        DateTime created = new(2024, 3, 17, 8, 30, 0, DateTimeKind.Utc);

        using MemoryStream stream = new();
        using (XLWorkbook workbook = new())
        {
            workbook.AddWorksheet("Data").Cell("A1").Value = "Hello";
            workbook.Properties.Author = "Ada";
            workbook.Properties.Title = "Quarterly";
            workbook.Properties.Category = "Reports";
            workbook.Properties.Created = created;
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);

        ClassicAssert.AreEqual("Ada", package.Properties.Creator);
        ClassicAssert.AreEqual("Quarterly", package.Properties.Title);
        ClassicAssert.AreEqual("Reports", package.Properties.Category);
        ClassicAssert.AreEqual(created, package.Properties.Created);
    }

    [Test]
    public void EditingPropertiesOfAnExistingPackageKeepsTheOtherOnes()
    {
        using MemoryStream original = new();
        using (XLWorkbook workbook = new())
        {
            workbook.AddWorksheet("Data").Cell("A1").Value = "Hello";
            workbook.Properties.Author = "Ada";
            workbook.Properties.Title = "Quarterly";
            workbook.SaveAs(original);
        }

        using MemoryStream rewritten = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, writable: true))
        {
            package.Properties.LastModifiedBy = "Grace";
            package.SaveTo(rewritten);
        }

        rewritten.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(rewritten);

        ClassicAssert.AreEqual("Grace", reopened.Properties.LastModifiedBy);
        ClassicAssert.AreEqual("Ada", reopened.Properties.Creator);
        ClassicAssert.AreEqual("Quarterly", reopened.Properties.Title);
    }

    [Test]
    public void AnUnrelatedPartUnderTheDefaultNameIsLeftAlone()
    {
        // A file whose core properties relationship type is misspelt still has its
        // /docProps/core.xml, just not related as the core properties.
        using MemoryStream original = new();
        using (OpcPackage package = OpcPackage.Create())
        {
            OpcPart stray = package.AddPart("/docProps/core.xml", OpcContentType.CoreProperties);
            using (StreamWriter writer = new(stray.GetWriteStream()))
            {
                writer.Write("<stray/>");
            }

            package.Relationships.Add(
                stray.Name,
                "http://schemas.openxmlformats.org/package/2006/relationships/meatadata/core-properties"
            );
            package.SaveTo(original);
        }

        using MemoryStream rewritten = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, writable: true))
        {
            package.Properties.Creator = "Ada";
            package.SaveTo(rewritten);
        }

        rewritten.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(rewritten);
        ClassicAssert.AreEqual("Ada", reopened.Properties.Creator);
        ClassicAssert.IsTrue(reopened.TryGetPart("/docProps/core.xml", out OpcPart? kept));
        using StreamReader reader = new(kept!.GetReadStream());
        ClassicAssert.AreEqual("<stray/>", reader.ReadToEnd());
    }

    private static OpcPart MinimalWorkbook(OpcPackage package)
    {
        OpcPart workbook = package.AddPart(
            "/xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        );

        using (Stream stream = workbook.GetWriteStream())
        using (StreamWriter writer = new(stream))
        {
            writer.Write(
                """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"""
            );
        }

        return workbook;
    }
}

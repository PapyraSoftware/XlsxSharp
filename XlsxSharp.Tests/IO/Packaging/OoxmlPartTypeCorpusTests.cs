using System.Reflection;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

/// <summary>
/// Checks <see cref="OoxmlPartTypes"/> against what Excel itself writes: in the Excel-authored
/// files of the test corpus, every part reached through one of the table's relationship types
/// has to be declared with the content type the table has for it.
/// </summary>
/// <remarks>
/// The files are the real reference - a wrong content type or relationship type in the table
/// would produce workbooks that differ from them, and Excel refuses to open a package whose
/// parts are declared with a content type it does not expect.
/// </remarks>
public class OoxmlPartTypeCorpusTests
{
    /// <summary>
    /// The part kinds with more than one content type: the four workbook variants share their
    /// relationship type and differ only in content type.
    /// </summary>
    private static readonly OoxmlPartType[] WorkbookVariants =
    [
        OoxmlPartTypes.Workbook,
        OoxmlPartTypes.MacroEnabledWorkbook,
        OoxmlPartTypes.WorkbookTemplate,
        OoxmlPartTypes.MacroEnabledWorkbookTemplate,
    ];

    /// <summary>
    /// Images have one relationship type but carry the content type of their format, so only the
    /// relationship type can be checked for them.
    /// </summary>
    private static readonly OoxmlPartType[] FixedContentTypes =
    [
        .. AllPartTypes().Where(t => !WorkbookVariants.Contains(t) && t != OoxmlPartTypes.Image),
    ];

    internal static IEnumerable<string> ExcelAuthoredFiles =>
        TestHelper.ListResourceFiles(s =>
            (s.Contains(".Schemas.Robustness.") || s.Contains(".Schemas.OpenXmlSdkCorpus."))
            // Strict files use a different relationship vocabulary altogether, see
            // RealWorldFixtureTests.KnownLimitations; the beta-era template predates the final one.
            && !s.Contains(".O14ISOStrict.")
            && !s.Contains("TestFiles.Comments.xlsx")
            && !s.Contains("ProjectStatusReport_TP10094814.xltx")
        );

    [Test]
    [MethodDataSource(nameof(ExcelAuthoredFiles))]
    public void ContentTypesMatchTheTable(string file)
    {
        using Stream resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(file));
        using MemoryStream stream = new();
        resource.CopyTo(stream);
        stream.Position = 0;

        using OpcPackage package = OpcPackage.Open(stream);

        foreach (OpcPart part in package.Parts)
        {
            // The package-level relationships lead to the workbook and the document properties.
            IEnumerable<OpcRelationship> incoming = package
                .Relationships.Concat(package.Parts.SelectMany(p => p.Relationships))
                .Where(r => OpcPartName.Comparer.Equals(r.TargetPartName, part.Name));

            foreach (OpcRelationship relationship in incoming)
            {
                AssertMatches(part, relationship.RelationshipType, file);
            }
        }
    }

    /// <summary>
    /// A kind of part the corpus never contains would not be checked by
    /// <see cref="ContentTypesMatchTheTable"/> at all.
    /// </summary>
    [Test]
    public void TheCorpusCoversEveryPartType()
    {
        HashSet<string> seen = [];
        foreach (string file in ExcelAuthoredFiles)
        {
            using Stream resource = TestHelper.GetStreamFromResource(
                TestHelper.GetResourcePath(file)
            );
            using MemoryStream stream = new();
            resource.CopyTo(stream);
            stream.Position = 0;

            using OpcPackage package = OpcPackage.Open(stream);
            foreach (
                OpcPart part in package.Parts.Where(p =>
                    AllPartTypes().Any(t => t.ContentType == p.ContentType)
                )
            )
            {
                seen.Add(part.ContentType);
            }
        }

        foreach (OoxmlPartType partType in AllPartTypes().Where(t => t != OoxmlPartTypes.Image))
        {
            ClassicAssert.IsTrue(
                seen.Contains(partType.ContentType),
                $"no file in the corpus has a {partType.ContentType} part"
            );
        }
    }

    private static void AssertMatches(OpcPart part, string relationshipType, string file)
    {
        if (WorkbookVariants.Any(t => t.RelationshipType == relationshipType))
        {
            // The officeDocument relationship is shared with other document kinds, so only a
            // SpreadsheetML workbook has to be one of the four.
            if (part.ContentType.Contains("sheet", StringComparison.OrdinalIgnoreCase))
            {
                ClassicAssert.IsTrue(
                    WorkbookVariants.Any(t => t.ContentType == part.ContentType),
                    $"{file}: {part.Name} is a workbook declared as {part.ContentType}"
                );
            }

            return;
        }

        if (OoxmlPartTypes.Image.RelationshipType == relationshipType)
        {
            ClassicAssert.IsTrue(
                part.ContentType.StartsWith("image/", StringComparison.Ordinal),
                $"{file}: {part.Name} is an image declared as {part.ContentType}"
            );
            return;
        }

        OoxmlPartType? expected = FixedContentTypes.FirstOrDefault(t =>
            t.RelationshipType == relationshipType
        );
        if (expected is not null)
        {
            ClassicAssert.AreEqual(
                expected.ContentType,
                part.ContentType,
                $"{file}: {part.Name}, reached through {relationshipType}"
            );
        }
    }

    private static IEnumerable<OoxmlPartType> AllPartTypes() =>
        typeof(OoxmlPartTypes)
            .GetProperties(BindingFlags.Public | BindingFlags.Static)
            .Where(p => p.PropertyType == typeof(OoxmlPartType))
            .Select(p => (OoxmlPartType)p.GetValue(null)!);
}

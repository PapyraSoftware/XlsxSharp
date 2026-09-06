using XlsxSharp.Excel;
using XlsxSharp.Excel.IO.Schemas;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.Excel.IO.Schemas;

/// <summary>
/// Exercises <see cref="SchemaValidator"/> and the general load/save round trip against real,
/// Excel-authored files taken from the Open XML SDK's own test suite (see THIRD-PARTY.txt) rather
/// than synthetic ones, the same way its "OFCAT"/Robustness corpus, its markup-compatibility
/// fixture and its missing-calcChain fixture are used there.
/// </summary>
public class RealWorldFixtureTests
{
    internal static IEnumerable<string> RobustnessFiles =>
        TestHelper.ListResourceFiles(s => s.Contains(".Schemas.Robustness."));

    /// <summary>
    /// Mirrors the Open XML SDK's own <c>Robustness.OFCATFull</c>: a corpus of real files found
    /// in the wild, each expected to validate cleanly and survive a load/save round trip.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(RobustnessFiles))]
    public void RealWorldFileValidatesAndRoundTrips(string file)
    {
        using MemoryStream original = LoadResource(file);

        using (OpcPackage package = OpcPackage.Open(original))
        {
            IReadOnlyList<string> errors = SchemaValidator.Validate(package);
            ClassicAssert.IsEmpty(errors, string.Join(Environment.NewLine, errors));
        }

        original.Position = 0;
        using XLWorkbook wb = new(original);
        using MemoryStream resaved = new();
        ClassicAssert.DoesNotThrow(() => wb.SaveAs(resaved));
    }

    /// <summary>
    /// Mirrors the SDK's <c>MCSupport.LoadProcessContent</c>: a shared string's <c>si</c> element
    /// carries <c>mc:Ignorable</c>/<c>mc:ProcessContent</c>/<c>mc:PreserveAttributes</c> content in
    /// a w14 extension namespace. XlsxSharp does not implement <c>mc:ProcessContent</c> - it only
    /// strips ignorable content wholesale - but the file still has to validate and round-trip.
    /// </summary>
    [Test]
    public void McExeclXlsxValidatesAndRoundTrips()
    {
        using MemoryStream original = LoadResource("Other.IO.Schemas.MCExecl.xlsx");

        using (OpcPackage package = OpcPackage.Open(original))
        {
            IReadOnlyList<string> errors = SchemaValidator.Validate(package);
            ClassicAssert.IsEmpty(errors, string.Join(Environment.NewLine, errors));
        }

        original.Position = 0;
        using XLWorkbook wb = new(original);
        ClassicAssert.AreEqual("abc", wb.Worksheets.First().Cell("A1").GetString());
        using MemoryStream resaved = new();
        ClassicAssert.DoesNotThrow(() => wb.SaveAs(resaved));
    }

    /// <summary>
    /// Mirrors the SDK's <c>M4Conformance.LoadExt2</c>: a real pivot chart's <c>extLst</c> nests
    /// an Office2010 extension several levels deep.
    /// </summary>
    [Test]
    public void ExtlstXlsxValidatesAndRoundTrips()
    {
        using MemoryStream original = LoadResource("Other.IO.Schemas.extlst.xlsx");

        using (OpcPackage package = OpcPackage.Open(original))
        {
            IReadOnlyList<string> errors = SchemaValidator.Validate(package);
            ClassicAssert.IsEmpty(errors, string.Join(Environment.NewLine, errors));
        }

        original.Position = 0;
        using XLWorkbook wb = new(original);
        using MemoryStream resaved = new();
        ClassicAssert.DoesNotThrow(() => wb.SaveAs(resaved));
    }

    /// <summary>
    /// Mirrors the SDK's own <c>OpenXmlPackageTests.SucceedWithMissingCalcChainPart</c>: the
    /// workbook part's relationships point at <c>/xl/calcChain.xml</c>, but the part itself was
    /// never actually written into the package. Unlike the SDK - which throws unless the caller
    /// opts into <c>IgnoreCalculationChainPartRelationship</c> - XlsxSharp treats calcChain as the
    /// purely advisory, Excel-regenerable part it is and tolerates the dangling reference.
    /// </summary>
    [Test]
    public void MissingCalcChainPartRoundTrips()
    {
        using MemoryStream original = LoadResource("Other.IO.Schemas.missingcalcchainpart.xlsx");

        using (OpcPackage package = OpcPackage.Open(original))
        {
            ClassicAssert.IsFalse(package.Parts.Any(p => p.Name == "/xl/calcChain.xml"));
        }

        original.Position = 0;
        using XLWorkbook wb = new(original);
        using MemoryStream resaved = new();
        ClassicAssert.DoesNotThrow(() => wb.SaveAs(resaved));
    }

    private static MemoryStream LoadResource(string resourcePath)
    {
        using Stream stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(resourcePath)
        );
        MemoryStream copy = new();
        stream.CopyTo(copy);
        copy.Position = 0;
        return copy;
    }
}

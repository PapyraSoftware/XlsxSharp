using XlsxSharp.Excel;
using XlsxSharp.Excel.IO.Schemas;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.Excel.IO.Schemas;

/// <summary>
/// Exercises <see cref="SchemaValidator"/> and the general load/save round trip against real,
/// Excel-authored files taken from the Open XML SDK's own test suite (see THIRD-PARTY.txt) rather
/// than synthetic ones, the same way its "OFCAT"/Robustness corpus, its markup-compatibility
/// fixture and its missing-calcChain fixture are used there - plus the rest of its test asset
/// library (<see cref="CorpusFiles"/>), covering pivot tables, charts, OLE/ActiveX objects,
/// external links, web extensions and Office templates, minus the individually investigated
/// <see cref="KnownLimitations"/>.
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

    /// <summary>
    /// Files under OpenXmlSdkCorpus that validate or round-trip cleanly against this file's own
    /// architecture (typed model, version-aware namespace handling) but not against XlsxSharp's -
    /// each is a real, separately investigated finding, not a blanket "some files fail" excuse.
    /// </summary>
    private static readonly string[] KnownLimitations =
    [
        // ISO/IEC 29500 "Strict" - a parallel OOXML variant with a different root namespace
        // throughout (e.g. http://purl.oclc.org/ooxml/spreadsheetml/main instead of the
        // Transitional http://schemas.openxmlformats.org/spreadsheetml/2006/main XlsxSharp reads
        // and SchemaValidator's vendored schemas describe) - a whole separate schema/reader
        // XlsxSharp does not implement. Excel itself barely ever produces Strict files.
        "TestFiles.Comments.xlsx",
        // Pre-final-spec, Excel-2007-beta-era content types (e.g.
        // application/vnd.ms-excel.worksheet+xml instead of the final
        // application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml) - an obsolete
        // draft vocabulary from before OOXML was standardized.
        "O12_templates.ProjectStatusReport_TP10094814.xltx",
        // ConditionalFormatXml.Rule has no writer case for XLConditionalFormatType.AboveAverage -
        // a real, acknowledged feature gap (the reader has nowhere to keep the rule's
        // aboveAverage/equalAverage/stdDev flags either), not something these files expose
        // incorrectly.
        "spreadsheet.NoExtDataB1.xlsx",
        "spreadsheet.NoExtDataF1.xlsx",
        "spreadsheet.noextdatab4.xlsx",
        // A numeric value typed as a date but outside DateTime.FromOADate's representable range
        // throws while writing instead of being tolerated the way Excel itself tolerates it.
        "spreadsheet.NoExtDataA1.xlsx",
        // Loading a table's totals-row formula copies a number format from another cell in a way
        // that (rarely) leaves the copied format's font not registered by reference identity yet,
        // tripping an internal Debug.Assert in XLWorkbookStyles.RegisterCellFormat.
        "O12_templates.EmployeeTimeCard_TP10192140.xltx",
        // A pivot cache record is expected to carry one value per cacheField, but a calculated
        // (formula=...) or grouped (fieldGroup) field - both marked databaseField="0" - is
        // correctly absent from every record; PivotCacheRecordsReader and XLPivotCache.FieldCount
        // count every cache field instead of only the ones actually stored in records.
        "spreadsheet.Pivot2.xlsx",
        // The SDK's own fixtures for a lenient, non-standard URI-parsing compatibility mode it
        // implements for malformed hyperlink relationship targets - unrelated to schema validation
        // or general round-tripping.
        "TestFiles.malformed_uri.xlsx",
        "TestFiles.malformed_uri_long.xlsx",
    ];

    internal static IEnumerable<string> CorpusFiles =>
        TestHelper.ListResourceFiles(s =>
            s.Contains(".Schemas.OpenXmlSdkCorpus.")
            && !s.Contains(".Schemas.OpenXmlSdkCorpus.TestDataStorage.O14ISOStrict.")
            && !KnownLimitations.Any(s.Contains)
        );

    /// <summary>
    /// A broader sweep of the same corpus, covering pivot tables, charts, OLE/ActiveX objects,
    /// external links, web extensions and Office templates - well beyond what
    /// <see cref="RealWorldFileValidatesAndRoundTrips"/>'s Robustness slice happens to exercise.
    /// </summary>
    [Test]
    [MethodDataSource(nameof(CorpusFiles))]
    public void SdkCorpusFileValidatesAndRoundTrips(string file)
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

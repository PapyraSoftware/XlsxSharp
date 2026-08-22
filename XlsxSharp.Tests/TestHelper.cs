using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;
using NUnit.Framework;
using XlsxSharp.Examples;
using XlsxSharp.Excel;
using XlsxSharp.Tests.Utils;
using LoadOptions = XlsxSharp.Excel.LoadOptions;
using Path = System.IO.Path;

namespace XlsxSharp.Tests;

internal static class TestHelper
{
    public static string CurrencySymbol =>
        Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol;

    //Note: Run example tests parameters
    public static string TestsOutputDirectory =>
        Path.Combine(
            Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
            "Generated"
        );

    public const string ActualTestResultPostFix = "";
    public static readonly string ExampleTestsOutputDirectory = Path.Combine(
        TestsOutputDirectory,
        "Examples"
    );

    private const bool CompareWithResources = true;

    private static readonly ResourceFileExtractor _extractor = new(
        Assembly.GetExecutingAssembly(),
        ".Resource."
    );

    public static void SaveWorkbook(XLWorkbook workbook, params string[] fileNameParts) =>
        workbook.SaveAs(
            Path.Combine(
                new string[] { TestsOutputDirectory }
                    .Concat(fileNameParts)
                    .ToArray()
            ),
            true
        );

    // Because different fonts are installed on Unix,
    // the columns widths after AdjustToContents() will
    // cause the tests to fail.
    // Therefore we ignore the width attribute when running on Unix
    public static bool StripColumnWidths => IsRunningOnUnix;

    public static bool IsRunningOnUnix
    {
        get
        {
            int p = (int)Environment.OSVersion.Platform;
            return ((p == 4) || (p == 6) || (p == 128));
        }
    }

    public static void RunTestExample<T>(string filePartName, bool evaluateFormulae = false)
        where T : IXLExample, new()
    {
        // Make sure tests run on a deterministic culture
        Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

        T example = new();
        string[] pathParts = filePartName.Split(new char[] { '\\' });
        string filePath1 = Path.Combine(
            new List<string>() { ExampleTestsOutputDirectory }
                .Concat(pathParts)
                .ToArray()
        );

        string extension = Path.GetExtension(filePath1);
        string? directory = Path.GetDirectoryName(filePath1);

        string fileName = Path.GetFileNameWithoutExtension(filePath1);
        fileName += ActualTestResultPostFix;
        fileName = Path.ChangeExtension(fileName, extension);

        filePath1 = Path.Combine(directory, "z" + fileName);
        string filePath2 = Path.Combine(directory, fileName);

        //Run test
        example.Create(filePath1);
        using (XLWorkbook wb = new(filePath1))
        {
            wb.SaveAs(filePath2, validate: true, evaluateFormulae);
        }

        // Also load from template and save it again - but not necessary to test against reference file
        // We're just testing that it can save.
        using (MemoryStream ms = new())
        using (XLWorkbook wb = XLWorkbook.OpenFromTemplate(filePath1))
        {
            wb.SaveAs(ms, validate: true, evaluateFormulae);
        }

        if (CompareWithResources)
        {
            string resourcePath = "Examples." + filePartName.Replace('\\', '.').TrimStart('.');
            using (Stream streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
            using (FileStream streamActual = File.OpenRead(filePath2))
            {
                bool success = ExcelDocsComparer.Compare(
                    streamActual,
                    streamExpected,
                    out string message
                );
                string formattedMessage = string.Format(
                    "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                    filePath2,
                    resourcePath,
                    message
                );

                Assert.IsTrue(success, formattedMessage);
            }
        }
    }

    /// <summary>
    /// Create a workbook and compare it with a saved resource.
    /// </summary>
    /// <param name="workbookGenerator">A function that gets an empty workbook and fills it with data.</param>
    /// <param name="referenceResource">Reference workbook saved in resources</param>
    /// <param name="evaluateFormulae">Should formulas of created workbook be evaluated and values saved?</param>
    /// <param name="validate">Should the created workbook be validated during by OpenXmlSdk validator?</param>
    public static void CreateAndCompare(
        Action<XLWorkbook> workbookGenerator,
        string referenceResource,
        bool evaluateFormulae = false,
        bool validate = true
    ) =>
        CreateAndCompare(
            () =>
            {
                XLWorkbook wb = new();
                workbookGenerator(wb);
                return wb;
            },
            referenceResource,
            evaluateFormulae,
            validate
        );

    public static void CreateAndCompare(
        Func<IXLWorkbook> workbookGenerator,
        string referenceResource,
        bool evaluateFormulae = false,
        bool validate = true
    )
    {
        Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

        string[] pathParts = referenceResource.Split(new char[] { '\\' });
        string filePath1 = Path.Combine(
            new List<string>() { TestsOutputDirectory }
                .Concat(pathParts)
                .ToArray()
        );

        string extension = Path.GetExtension(filePath1);
        string? directory = Path.GetDirectoryName(filePath1);

        string fileName = Path.GetFileNameWithoutExtension(filePath1);
        fileName += ActualTestResultPostFix;
        fileName = Path.ChangeExtension(fileName, extension);

        string filePath2 = Path.Combine(directory, fileName);

        using (IXLWorkbook wb = workbookGenerator.Invoke())
        {
            wb.SaveAs(filePath2, validate, evaluateFormulae);
        }

        if (CompareWithResources)
        {
            string resourcePath = referenceResource.Replace('\\', '.').TrimStart('.');
            using (Stream streamExpected = _extractor.ReadFileFromResourceToStream(resourcePath))
            using (FileStream streamActual = File.OpenRead(filePath2))
            {
                bool success = ExcelDocsComparer.Compare(
                    streamActual,
                    streamExpected,
                    out string message
                );
                string formattedMessage = string.Format(
                    "Actual file '{0}' is different than the expected file '{1}'. The difference is: '{2}'",
                    filePath2,
                    resourcePath,
                    message
                );

                Assert.IsTrue(success, formattedMessage);
            }
        }
    }

    /// <summary>
    /// Load a file from the <paramref name="loadResourcePath"/>, modify it, save it through XlsxSharp
    /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
    /// </summary>
    /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
    public static void LoadModifyAndCompare(
        string loadResourcePath,
        Action<XLWorkbook> modify,
        string expectedOutputResourcePath,
        bool evaluateFormulae = false,
        bool validate = true
    )
    {
        using Stream stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
        using MemoryStream ms = new();
        CreateAndCompare(
            () =>
            {
                XLWorkbook wb = new(stream);
                modify(wb);
                return wb;
            },
            expectedOutputResourcePath,
            evaluateFormulae,
            validate
        );
    }

    /// <summary>
    /// Load a file from the <paramref name="loadResourcePath"/>, save it through XlsxSharp without modifications
    /// and compare the saved file against the <paramref name="expectedOutputResourcePath"/>.
    /// </summary>
    /// <remarks>Useful for checking whether we can load data from Excel and save it while keeping various feature in the OpenXML intact.</remarks>
    public static void LoadSaveAndCompare(
        string loadResourcePath,
        string expectedOutputResourcePath,
        bool evaluateFormulae = false,
        bool validate = true
    ) =>
        LoadModifyAndCompare(
            loadResourcePath,
            _ => { },
            expectedOutputResourcePath,
            evaluateFormulae,
            validate
        );

    /// <summary>
    /// A testing method to load a workbook from resource and assert the state of the loaded workbook.
    /// </summary>
    public static void LoadAndAssert(
        Action<XLWorkbook> assertWorkbook,
        string loadResourcePath,
        LoadOptions options = null
    )
    {
        using Stream stream = GetStreamFromResource(GetResourcePath(loadResourcePath));
        using XLWorkbook wb = new(stream, options ?? new LoadOptions());

        assertWorkbook(wb);
    }

    /// <summary>
    /// A testing method to load a workbook with a single worksheet from resource and assert
    /// the state of the loaded workbook.
    /// </summary>
    public static void LoadAndAssert(
        Action<XLWorkbook, IXLWorksheet> assertWorksheet,
        string loadResourcePath,
        LoadOptions options = null
    ) =>
        LoadAndAssert(
            wb =>
            {
                IXLWorksheet ws = wb.Worksheets.Single();
                assertWorksheet(wb, ws);
            },
            loadResourcePath,
            options
        );

    public static string GetResourcePath(string filePartName) =>
        filePartName.Replace('\\', '.').TrimStart('.');

    public static Stream GetStreamFromResource(string resourcePath) =>
        _extractor.ReadFileFromResourceToStream(resourcePath);

    public static void LoadFile(string filePartName)
    {
        IXLWorkbook wb;
        using (Stream stream = GetStreamFromResource(GetResourcePath(filePartName)))
        {
            Assert.DoesNotThrow(
                () => wb = new XLWorkbook(stream),
                "Unable to load resource {0}",
                filePartName
            );
        }
    }

    public static IEnumerable<string> ListResourceFiles(Func<string, bool> predicate = null) =>
        _extractor.GetFileNames(predicate);

    /// <summary>
    /// A method for testing of a saving and loading capabilities of XlsxSharp. Use this
    /// method to check properties are correctly saved and loaded.
    /// </summary>
    /// <remarks>This method is specialized, so it only works on one sheet.</remarks>
    /// <param name="createWorksheet">
    /// Method to setup a worksheet that will be saved and the saved file will be compared to
    /// <paramref name="referenceResource"/>.
    /// </param>
    /// <param name="assertLoadedWorkbook">
    /// <paramref name="referenceResource"/> will be loaded and this method will check that it
    /// was loaded correctly (i.e. properties are what was set in <paramref name="createWorksheet"/>).
    /// </param>
    /// <param name="referenceResource">Saved reference file.</param>
    public static void CreateSaveLoadAssert(
        Action<XLWorkbook, IXLWorksheet> createWorksheet,
        Action<XLWorkbook, IXLWorksheet> assertLoadedWorkbook,
        string referenceResource
    )
    {
        CreateAndCompare(
            wb =>
            {
                IXLWorksheet ws = wb.AddWorksheet();
                createWorksheet(wb, ws);
            },
            referenceResource
        );
        LoadAndAssert(assertLoadedWorkbook, referenceResource);
    }

    /// <summary>
    /// Basically can survive through save and load cycle. Doesn't check against actual file.
    /// Useful for testing is internal structures are correctly initialized after load.
    /// </summary>
    /// <param name="createWorksheet">Code to create a workbook.</param>
    /// <param name="assertLoadedWorkbook">Method to assert that workbook was loaded correctly.</param>
    /// <param name="validate">Validate created workbook that it is a valid OOXML file by OpenXML SDK.</param>
    /// <param name="evaluateFormulas">Evaluate formulas during saving and save the evaluated results to the workbook file.</param>
    public static void CreateSaveLoadAssert(
        Action<XLWorkbook, IXLWorksheet> createWorksheet,
        Action<XLWorkbook, IXLWorksheet> assertLoadedWorkbook,
        bool validate = true,
        bool evaluateFormulas = false
    )
    {
        using MemoryStream ms = new();
        using (XLWorkbook wb = new())
        {
            IXLWorksheet ws = wb.AddWorksheet();
            createWorksheet(wb, ws);
            wb.SaveAs(ms, validate, evaluateFormulas);
        }

        using (XLWorkbook wb = new(ms))
        {
            IXLWorksheet ws = wb.Worksheets.Single();
            assertLoadedWorkbook(wb, ws);
        }
    }

    /// <summary>
    /// Test if some aspect of a workbook can survive through save and load cycle.
    /// </summary>
    public static void CreateSaveLoadAssert(
        Action<XLWorkbook> createWorksheet,
        Action<XLWorkbook> assertLoadedWorkbook,
        bool validate = true,
        bool evaluateFormulas = false
    )
    {
        using MemoryStream ms = new();
        using (XLWorkbook wb = new())
        {
            createWorksheet(wb);
            wb.SaveAs(ms, validate, evaluateFormulas);
        }

        ms.Position = 0;
        using (XLWorkbook wb = new(ms))
        {
            assertLoadedWorkbook(wb);
        }
    }

    /// <summary>
    /// A method for testing the saving of a workbook. It loads and saves the workbook and uses
    /// assert methods for individual parts to check that the relevant parts were saved correctly.
    /// Unlike the <see cref="LoadSaveAndCompare"/>, this method is more resistant to changes
    /// in the saving code that are not directly related to the tested parts.
    /// </summary>
    internal static void LoadSaveAndAssert(
        string referenceResource,
        string part1,
        Action<XDocument> part1Assert,
        string part2,
        Action<XDocument> part2Assert
    ) => LoadSaveAndAssert(referenceResource, [(part1, part1Assert), (part2, part2Assert)]);

    private static void LoadSaveAndAssert(
        string loadResourcePath,
        (string PartPath, Action<XDocument> PartAssert)[] parts
    )
    {
        using MemoryStream ms = new();
        using (Stream stream = GetStreamFromResource(GetResourcePath(loadResourcePath)))
        {
            using XLWorkbook wb = new(stream);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using Package package = Package.Open(ms, FileMode.Open, FileAccess.Read);
        foreach ((string partPath, Action<XDocument> partAssert) in parts)
        {
            PackagePart part = package.GetPart(new Uri(partPath, UriKind.Relative));
            using Stream partStream = part.GetStream();
            XDocument partXml = XDocument.Load(partStream);
            partAssert(partXml);
        }
    }
}

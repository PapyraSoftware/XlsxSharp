using System.Xml.Linq;
using XlsxSharp.Excel;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.Excel.IO;

/// <summary>
/// Covers the guarantee <see cref="XlsxSharp.Excel.IO.WorksheetPartWriter"/>'s streaming worksheet
/// reader relies on: everything outside <c>sheetData</c> is carried through exactly as loaded,
/// even content XlsxSharp itself does not recognise, while <c>sheetData</c> itself is always
/// rebuilt from the workbook model regardless of what the loaded part held.
/// </summary>
public class WorksheetPartWriterTests
{
    private const string WorksheetPartName = "/xl/worksheets/sheet1.xml";
    private const string MainNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string ForeignNamespace = "urn:test:foreign";

    [Test]
    public void ForeignContentOutsideSheetDataSurvivesAResave()
    {
        using MemoryStream original = new();
        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = "before";
            wb.SaveAs(original);
        }

        // Inject an element XlsxSharp has no concept of into the worksheet part's extLst, the way
        // a newer Excel feature this version of the schema does not know about would show up -
        // extLst/ext is the one place CT_Worksheet actually allows unrecognised content; a bare
        // unknown element directly under <worksheet> would not be valid OOXML in the first place.
        using MemoryStream withForeignContent = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, withForeignContent))
        {
            OpcPart worksheetPart = package.GetPart(WorksheetPartName);
            XDocument document = Load(worksheetPart);
            document.Root!.Add(
                new XElement(
                    XNamespace.Get(MainNamespace) + "extLst",
                    new XElement(
                        XNamespace.Get(MainNamespace) + "ext",
                        new XAttribute("uri", "{12345678-1234-1234-1234-123456789012}"),
                        new XElement(XNamespace.Get(ForeignNamespace) + "custom", "hello")
                    )
                )
            );
            Save(worksheetPart, document);
        }

        // Load that file, change a cell through the model, and save again.
        using MemoryStream resaved = new();
        withForeignContent.Position = 0;
        using (XLWorkbook wb = new(withForeignContent))
        {
            wb.Worksheet("Sheet1").Cell("A1").Value = "after";
            wb.SaveAs(resaved);
        }

        // The foreign element is still there, untouched by never having been understood.
        resaved.Position = 0;
        using (OpcPackage reopened = OpcPackage.Open(resaved))
        {
            XElement? foreignElement = Load(reopened.GetPart(WorksheetPartName))
                .Root!.Element(XNamespace.Get(MainNamespace) + "extLst")
                ?.Element(XNamespace.Get(MainNamespace) + "ext")
                ?.Element(XNamespace.Get(ForeignNamespace) + "custom");
            ClassicAssert.IsNotNull(foreignElement);
            ClassicAssert.AreEqual("hello", foreignElement!.Value);
        }

        // sheetData itself was rebuilt from the model rather than carried through - the row still
        // reflects the change made above, not what the loaded part held.
        resaved.Position = 0;
        using XLWorkbook reloaded = new(resaved);
        ClassicAssert.AreEqual("after", reloaded.Worksheet("Sheet1").Cell("A1").GetString());
    }

    private static XDocument Load(OpcPart part)
    {
        using Stream stream = part.GetReadStream();
        return XDocument.Load(stream);
    }

    private static void Save(OpcPart part, XDocument document)
    {
        using Stream stream = part.GetWriteStream();
        document.Save(stream);
    }
}

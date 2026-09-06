using XlsxSharp.Excel;
using XlsxSharp.Excel.IO.Schemas;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.Excel.IO.Schemas;

public class SchemaValidatorTests
{
    [Test]
    public void AValidPackageHasNoErrors()
    {
        using MemoryStream stream = new();
        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = "hello";
            wb.SaveAs(stream, validate: false);
        }

        stream.Position = 0;
        using OpcPackage package = OpcPackage.Open(stream);
        ClassicAssert.IsEmpty(SchemaValidator.Validate(package));
    }

    [Test]
    public void AnInvalidChildElementIsReported()
    {
        const string worksheetPartName = "/xl/worksheets/sheet1.xml";

        using MemoryStream original = new();
        using (XLWorkbook wb = new())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = "hello";
            wb.SaveAs(original, validate: false);
        }

        using MemoryStream corrupted = new();
        original.Position = 0;
        using (OpcPackage package = OpcPackage.Open(original, corrupted))
        {
            OpcPart worksheetPart = package.GetPart(worksheetPartName);

            // "definedNames" is a real CT_Worksheet element, but it belongs under <workbook>,
            // not <worksheet> - a schema-invalid element in a place that is not a wildcard/
            // extension point, which the validator has to catch on its own rather than have
            // XlsxSharp's own writer ever have produced.
            using Stream writeStream = worksheetPart.GetWriteStream();
            writeStream.Write(
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><definedNames/></worksheet>"u8
            );
        }

        corrupted.Position = 0;
        using OpcPackage reopened = OpcPackage.Open(corrupted);
        IReadOnlyList<string> errors = SchemaValidator.Validate(reopened);

        ClassicAssert.IsTrue(
            errors.Count > 0,
            "Expected the invalid worksheet part to be reported."
        );
        ClassicAssert.IsTrue(
            errors.Any(e => e.Contains(worksheetPartName, StringComparison.Ordinal))
        );
    }
}

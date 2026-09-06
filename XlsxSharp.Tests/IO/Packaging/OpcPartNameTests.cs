using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Tests.IO.Packaging;

public class OpcPartNameTests
{
    [Test]
    [Arguments("/xl/workbook.xml", "/xl/workbook.xml")]
    [Arguments("xl/workbook.xml", "/xl/workbook.xml")]
    [Arguments("/[Content_Types].xml", "/[Content_Types].xml")]
    public void NormalizeAddsTheLeadingSlash(string input, string expected) =>
        ClassicAssert.AreEqual(expected, OpcPartName.Normalize(input));

    [Test]
    [Arguments("")]
    [Arguments("/")]
    [Arguments("/xl/")]
    [Arguments("/xl//workbook.xml")]
    [Arguments("/xl/./workbook.xml")]
    [Arguments("/xl/../workbook.xml")]
    [Arguments("/xl/workbook.")]
    public void NormalizeRejectsInvalidNames(string input) =>
        ClassicAssert.Throws<OpcException>(() => OpcPartName.Normalize(input));

    [Test]
    [Arguments("/xl/workbook.xml", "worksheets/sheet1.xml", "/xl/worksheets/sheet1.xml")]
    [Arguments("/xl/workbook.xml", "styles.xml", "/xl/styles.xml")]
    [Arguments("/xl/workbook.xml", "/xl/styles.xml", "/xl/styles.xml")]
    [Arguments(
        "/xl/worksheets/sheet1.xml",
        "../drawings/drawing1.xml",
        "/xl/drawings/drawing1.xml"
    )]
    [Arguments("/xl/worksheets/sheet1.xml", "./sheet2.xml", "/xl/worksheets/sheet2.xml")]
    [Arguments("", "docProps/core.xml", "/docProps/core.xml")]
    [Arguments("", "xl/workbook.xml", "/xl/workbook.xml")]
    public void ResolveTargetIsRelativeToTheSourceFolder(
        string source,
        string target,
        string expected
    ) => ClassicAssert.AreEqual(expected, OpcPartName.ResolveTarget(source, target));

    [Test]
    public void ResolveTargetRejectsEscapingThePackageRoot() =>
        ClassicAssert.Throws<OpcException>(() =>
            OpcPartName.ResolveTarget("/xl/workbook.xml", "../../outside.xml")
        );

    [Test]
    [Arguments("/xl/workbook.xml", "/xl/worksheets/sheet1.xml", "worksheets/sheet1.xml")]
    [Arguments("/xl/workbook.xml", "/xl/styles.xml", "styles.xml")]
    [Arguments(
        "/xl/worksheets/sheet1.xml",
        "/xl/drawings/drawing1.xml",
        "../drawings/drawing1.xml"
    )]
    [Arguments("", "/docProps/core.xml", "docProps/core.xml")]
    [Arguments("", "/xl/workbook.xml", "xl/workbook.xml")]
    public void MakeRelativeTargetIsTheInverseOfResolveTarget(
        string source,
        string target,
        string expected
    )
    {
        string relative = OpcPartName.MakeRelativeTarget(source, target);
        ClassicAssert.AreEqual(expected, relative);
        ClassicAssert.AreEqual(target, OpcPartName.ResolveTarget(source, relative));
    }

    [Test]
    [Arguments("", "/_rels/.rels")]
    [Arguments("/xl/workbook.xml", "/xl/_rels/workbook.xml.rels")]
    [Arguments("/xl/worksheets/sheet1.xml", "/xl/worksheets/_rels/sheet1.xml.rels")]
    public void GetRelationshipPartNameSitsInTheRelsFolder(string partName, string expected) =>
        ClassicAssert.AreEqual(expected, OpcPartName.GetRelationshipPartName(partName));

    [Test]
    [Arguments("/_rels/.rels", true)]
    [Arguments("/xl/_rels/workbook.xml.rels", true)]
    [Arguments("/xl/workbook.xml", false)]
    [Arguments("/xl/rels.xml", false)]
    public void IsRelationshipPartLooksAtTheRelsFolder(string partName, bool expected) =>
        ClassicAssert.AreEqual(expected, OpcPartName.IsRelationshipPart(partName));

    [Test]
    [Arguments("/xl/workbook.xml", "xml")]
    [Arguments("/xl/media/image1.PNG", "png")]
    [Arguments("/_rels/.rels", "rels")]
    [Arguments("/xl/noextension", "")]
    public void GetExtensionIsLowerCasedAndWithoutTheDot(string partName, string expected) =>
        ClassicAssert.AreEqual(expected, OpcPartName.GetExtension(partName));
}

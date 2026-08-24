using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

public class PointTests
{
    [Test]
    [Arguments("A1", 1, 1)]
    [Arguments("AA1", 27, 1)]
    [Arguments("AAA1", 703, 1)]
    [Arguments("Z1", 26, 1)]
    [Arguments("ZZ1", 702, 1)]
    [Arguments("XFD1", 16384, 1)]
    [Arguments("A1", 1, 1)]
    [Arguments("A999", 1, 999)]
    [Arguments("XFD1048576", 16384, 1048576)]
    public void ParseCellRefsAccordingToGrammar(string cellRef, int columnNumber, int rowNumber)
    {
        Point sheetPoint = Point.Parse(cellRef.AsSpan());
        ClassicAssert.AreEqual(columnNumber, sheetPoint.Column);
        ClassicAssert.AreEqual(rowNumber, sheetPoint.Row);
    }

    [Test]
    [Arguments("")]
    [Arguments(" ")]
    [Arguments("A")]
    [Arguments("AA")]
    [Arguments("1")]
    [Arguments("11")]
    [Arguments(" A1")]
    [Arguments("A1 ")]
    [Arguments("A 1")]
    [Arguments("@1")] // @ is a char 'A' - 1
    [Arguments("[1")] // [ is a char 'Z' + 1
    [Arguments("A:")] // : is a char '9' + 1
    [Arguments("A/")] // / is a char '0' - 1
    [Arguments("A1:")]
    [Arguments("A1/")]
    [Arguments("A@1")]
    [Arguments("A[1")]
    [Arguments("XFE1")]
    [Arguments("AAAA1")]
    [Arguments("A1048577")]
    [Arguments("A01")]
    [Arguments("A0")]
    [Arguments("A-1")]
    public void InvalidInputsAreNotParsed(string cellRef) =>
        ClassicAssert.Throws<FormatException>(() => Point.Parse(cellRef.AsSpan()));

    [Test]
    [Arguments("A1")]
    [Arguments("DE1")]
    [Arguments("D174")]
    [Arguments("XFD1048576")]
    public void CanFormatToString(string cellRef)
    {
        Point r = Point.Parse(cellRef);
        ClassicAssert.AreEqual(cellRef, r.ToString());
    }
}

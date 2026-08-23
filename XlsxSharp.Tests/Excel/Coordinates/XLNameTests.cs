using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

public class XlNameTests
{
    [Test]
    public void WorkbookScopedNameIsComparedCaseInsensitive()
    {
        XLName lowerCase = new("name");
        XLName upperCase = new("NAME");

        ClassicAssert.AreEqual(lowerCase, upperCase);
        ClassicAssert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

        ClassicAssert.AreNotEqual(lowerCase, new XLName("different_name"));
    }

    [Test]
    public void SheetScopedNameIsComparedCaseInsensitive()
    {
        XLName lowerCase = new("sheet", "name");
        XLName upperCase = new("SHEET", "NAME");

        ClassicAssert.AreEqual(lowerCase, upperCase);
        ClassicAssert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

        ClassicAssert.AreNotEqual(lowerCase, new XLName("Different sheet", "name"));
        ClassicAssert.AreNotEqual(lowerCase, new XLName("sheet", "different_name"));
    }
}

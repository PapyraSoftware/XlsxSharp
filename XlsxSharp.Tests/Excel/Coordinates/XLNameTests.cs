using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

[TestFixture]
public class XLNameTests
{
    [Test]
    public void WorkbookScopedNameIsComparedCaseInsensitive()
    {
        XLName lowerCase = new("name");
        XLName upperCase = new("NAME");

        Assert.AreEqual(lowerCase, upperCase);
        Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

        Assert.AreNotEqual(lowerCase, new XLName("different_name"));
    }

    [Test]
    public void SheetScopedNameIsComparedCaseInsensitive()
    {
        XLName lowerCase = new("sheet", "name");
        XLName upperCase = new("SHEET", "NAME");

        Assert.AreEqual(lowerCase, upperCase);
        Assert.AreEqual(lowerCase.GetHashCode(), upperCase.GetHashCode());

        Assert.AreNotEqual(lowerCase, new XLName("Different sheet", "name"));
        Assert.AreNotEqual(lowerCase, new XLName("sheet", "different_name"));
    }
}

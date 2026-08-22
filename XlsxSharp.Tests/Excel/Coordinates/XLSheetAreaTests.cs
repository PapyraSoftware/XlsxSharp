using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

[TestFixture]
public class XLSheetAreaTests
{
    [Test]
    public void SheetNameIsComparedCaseInsensitive()
    {
        SheetArea upperCase = new("NAME", new Area(1, 2, 3, 4));
        SheetArea lowerCase = new("name", new Area(1, 2, 3, 4));
        Assert.AreEqual(upperCase.GetHashCode(), lowerCase.GetHashCode());
        Assert.AreEqual(upperCase, lowerCase);
    }

    [Test]
    public void IntersectionProducesRangeIntersectionInSameSheet()
    {
        SheetArea sheetArea1 = new("SHEET", Area.Parse("A1:C3"));
        SheetArea sheetArea2 = new("sheet", Area.Parse("B2:D4"));
        SheetArea otherSheetArea = new("Other", Area.Parse("B2:D4"));

        SheetArea? sameSheetIntersection = sheetArea1.Intersect(sheetArea2);
        Assert.AreEqual(new SheetArea("sheet", Area.Parse("B2:C3")), sameSheetIntersection);

        SheetArea? differentSheetIntersection = sheetArea1.Intersect(otherSheetArea);
        Assert.Null(differentSheetIntersection);
    }
}

using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

[TestFixture]
public class XLSheetAreaTests
{
    [Test]
    public void Sheet_name_is_compared_case_insensitive()
    {
        SheetArea upperCase = new("NAME", new Area(1, 2, 3, 4));
        SheetArea lowerCase = new("name", new Area(1, 2, 3, 4));
        Assert.AreEqual(upperCase.GetHashCode(), lowerCase.GetHashCode());
        Assert.AreEqual(upperCase, lowerCase);
    }

    [Test]
    public void Intersection_produces_range_intersection_in_same_sheet()
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

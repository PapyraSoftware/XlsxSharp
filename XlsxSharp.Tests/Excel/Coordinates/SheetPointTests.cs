using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

internal class SheetPointTests
{
    [Test]
    public void Ctor_sheet_must_not_be_null()
    {
        ArgumentNullException ex = ClassicAssert.Throws<ArgumentNullException>(() =>
            new SheetPoint(null, new Point(1, 1))
        );
        ClassicAssert.AreEqual("sheetName", ex.ParamName);
    }

    [Test]
    public void Ctor_sheet_must_not_be_empty()
    {
        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() =>
            new SheetPoint("", new Point(1, 1))
        );
        ClassicAssert.AreEqual("sheetName", ex.ParamName);
    }

    [Test]
    [Arguments("sheet", 2, 5, "sheet", 2, 5, true)]
    [Arguments("SHEET", 2, 5, "sheet", 2, 5, true)]
    [Arguments("sheet", 2, 5, "sheet", 2, 6, false)]
    [Arguments("SHEET", 2, 5, "sheet", 3, 5, false)]
    [Arguments("some sheet", 2, 5, "other sheet", 2, 5, false)]
    public void Two_points_are_compared_by_case_insensitive_sheet_name_and_point_coordinates(
        string firstName,
        int firstRow,
        int firstColumn,
        string secondName,
        int secondRow,
        int secondColumn,
        bool areEqual
    )
    {
        SheetPoint first = new(firstName, firstRow, firstColumn);
        SheetPoint second = new(secondName, secondRow, secondColumn);
        ClassicAssert.AreEqual(areEqual, first == second);
        ClassicAssert.AreEqual(areEqual, first.GetHashCode() == second.GetHashCode());
    }

    [Test]
    [Arguments("sheet", 1, 4, "sheet!D1")]
    [Arguments("Joe's", 47, 28, "'Joe''s'!AB47")]
    [Arguments("2025 Q1", XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber, "'2025 Q1'!XFD1048576")]
    public void ToString_returns_readable_reference(
        string name,
        int row,
        int column,
        string expected
    )
    {
        SheetPoint bookPoint = new(name, row, column);
        ClassicAssert.AreEqual(expected, bookPoint.ToString());
    }
}

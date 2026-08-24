using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

internal class XlRowAreaTests
{
    [Test]
    public void Ctor_sheet_must_not_be_null()
    {
        ArgumentNullException ex = ClassicAssert.Throws<ArgumentNullException>(() =>
            new XLRowArea(null, 1)
        );
        ClassicAssert.AreEqual("name", ex.ParamName);
    }

    [Test]
    public void Ctor_sheet_must_not_be_empty()
    {
        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() => new XLRowArea("", 1));
        ClassicAssert.AreEqual("name", ex.ParamName);
    }

    [Test]
    [Arguments(-50)]
    [Arguments(0)]
    [Arguments(XLHelper.MaxRowNumber + 1)]
    [Arguments(int.MaxValue)]
    public void Ctor_row_number_must_be_valid(int invalidRowNumber) =>
        ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
            new XLRowArea("some sheet", invalidRowNumber)
        );

    [Test]
    [Arguments("name", 5, "name", 5, true)]
    [Arguments("NAME", 5, "name", 5, true)]
    [Arguments("NAME", 5, "name", 4, false)]
    [Arguments("some name", 1, "other name", 1, false)]
    public void Two_areas_are_compared_by_case_insensitive_sheet_name_and_row_number(
        string firstName,
        int firstRow,
        string secondName,
        int secondRow,
        bool areEqual
    )
    {
        XLRowArea first = new(firstName, firstRow);
        XLRowArea second = new(secondName, secondRow);
        ClassicAssert.AreEqual(areEqual, first == second);
        ClassicAssert.AreEqual(areEqual, first.GetHashCode() == second.GetHashCode());
    }

    [Test]
    public void Area_property_returns_area_of_row()
    {
        XLRowArea row = new("name", 4);
        SheetArea rowArea = row.Area;
        ClassicAssert.AreEqual(
            rowArea,
            new SheetArea(
                "name",
                new Area(4, XLHelper.MinColumnNumber, 4, XLHelper.MaxColumnNumber)
            )
        );
    }

    [Test]
    [Arguments("name", 4, "name!4:4")]
    [Arguments("some name", 4, "'some name'!4:4")]
    [Arguments("Joe's", 4, "'Joe''s'!4:4")]
    [Arguments("Joe", XLHelper.MaxRowNumber, "Joe!1048576:1048576")]
    public void ToString_returns_readable_reference(string name, int rowNumber, string expected)
    {
        XLRowArea rowArea = new(name, rowNumber);
        ClassicAssert.AreEqual(expected, rowArea.ToString());
    }
}

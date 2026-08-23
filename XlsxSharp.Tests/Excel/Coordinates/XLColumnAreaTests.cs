using System;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

internal class XlColumnAreaTests
{
    [Test]
    public void Ctor_sheet_must_not_be_null()
    {
        ArgumentNullException ex = ClassicAssert.Throws<ArgumentNullException>(() =>
            new XLColumnArea(null, 1)
        );
        ClassicAssert.AreEqual("name", ex.ParamName);
    }

    [Test]
    public void Ctor_sheet_must_not_be_empty()
    {
        ArgumentException ex = ClassicAssert.Throws<ArgumentException>(() =>
            new XLColumnArea("", 1)
        );
        ClassicAssert.AreEqual("name", ex.ParamName);
    }

    [Test]
    [Arguments(-50)]
    [Arguments(0)]
    [Arguments(XLHelper.MaxColumnNumber + 1)]
    [Arguments(int.MaxValue)]
    public void Ctor_column_number_must_be_valid(int invalidColumnNumber) =>
        ClassicAssert.Throws<ArgumentOutOfRangeException>(() =>
            new XLColumnArea("some sheet", invalidColumnNumber)
        );

    [Test]
    [Arguments("name", 5, "name", 5, true)]
    [Arguments("NAME", 5, "name", 5, true)]
    [Arguments("NAME", 5, "name", 4, false)]
    [Arguments("some name", 1, "other name", 1, false)]
    public void Two_areas_are_compared_by_case_insensitive_sheet_name_and_column_number(
        string firstName,
        int firstColumn,
        string secondName,
        int secondColumn,
        bool areEqual
    )
    {
        XLColumnArea first = new(firstName, firstColumn);
        XLColumnArea second = new(secondName, secondColumn);
        ClassicAssert.AreEqual(areEqual, first == second);
        ClassicAssert.AreEqual(areEqual, first.GetHashCode() == second.GetHashCode());
    }

    [Test]
    public void Area_property_returns_area_of_column()
    {
        XLColumnArea column = new("name", 4);
        SheetArea columnArea = column.Area;
        ClassicAssert.AreEqual(
            columnArea,
            new SheetArea("name", new Area(1, 4, XLHelper.MaxRowNumber, 4))
        );
    }

    [Test]
    [Arguments("name", 4, "name!D:D")]
    [Arguments("some name", 26, "'some name'!Z:Z")]
    [Arguments("Joe's", 27, "'Joe''s'!AA:AA")]
    [Arguments("Joe", XLHelper.MaxColumnNumber, "Joe!XFD:XFD")]
    public void ToString_returns_readable_reference(string name, int columnNumber, string expected)
    {
        XLColumnArea columnArea = new(name, columnNumber);
        ClassicAssert.AreEqual(expected, columnArea.ToString());
    }
}

using System;
using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

[TestOf(typeof(XLRowArea))]
internal class XlRowAreaTests
{
    [Test]
    public void Ctor_sheet_must_not_be_null() =>
        Assert.That(
            () => new XLRowArea(null, 1),
            Throws
                .Exception.TypeOf<ArgumentNullException>()
                .With.Property("ParamName")
                .EqualTo("name")
        );

    [Test]
    public void Ctor_sheet_must_not_be_empty() =>
        Assert.That(
            () => new XLRowArea("", 1),
            Throws.Exception.TypeOf<ArgumentException>().With.Property("ParamName").EqualTo("name")
        );

    [TestCase(-50)]
    [TestCase(0)]
    [TestCase(XLHelper.MaxRowNumber + 1)]
    [TestCase(int.MaxValue)]
    public void Ctor_row_number_must_be_valid(int invalidRowNumber) =>
        Assert.That(
            () => new XLRowArea("some sheet", invalidRowNumber),
            Throws.Exception.TypeOf<ArgumentOutOfRangeException>()
        );

    [TestCase("name", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 5, true)]
    [TestCase("NAME", 5, "name", 4, false)]
    [TestCase("some name", 1, "other name", 1, false)]
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
        Assert.That(first == second, Is.EqualTo(areEqual));
        Assert.That(first.GetHashCode() == second.GetHashCode(), Is.EqualTo(areEqual));
    }

    [Test]
    public void Area_property_returns_area_of_row()
    {
        XLRowArea row = new("name", 4);
        SheetArea rowArea = row.Area;
        Assert.AreEqual(
            rowArea,
            new SheetArea(
                "name",
                new Area(4, XLHelper.MinColumnNumber, 4, XLHelper.MaxColumnNumber)
            )
        );
    }

    [TestCase("name", 4, "name!4:4")]
    [TestCase("some name", 4, "'some name'!4:4")]
    [TestCase("Joe's", 4, "'Joe''s'!4:4")]
    [TestCase("Joe", XLHelper.MaxRowNumber, "Joe!1048576:1048576")]
    public void ToString_returns_readable_reference(string name, int rowNumber, string expected)
    {
        XLRowArea rowArea = new(name, rowNumber);
        Assert.AreEqual(expected, rowArea.ToString());
    }
}

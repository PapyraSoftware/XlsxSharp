using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

internal class XlAreaListTests
{
    [Test]
    [Arguments("A1:C3", "A1", "B1:C3 A2:A4")]
    [Arguments("A1:C3", "B1", "A1:A3 C1:C3 B2:B4")]
    [Arguments("A1:C3", "C1", "A1:B3 C2:C4")]
    [Arguments("A1:C3", "A2", "A1:C1 B2:C3 A2:A4")]
    [Arguments("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2:B4")]
    [Arguments("A1:C3", "C2", "A1:C1 A2:B3 C2:C4")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:C3 A3:A4")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 C3 B3:B4")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3 C3:C4")]
    [Arguments("B1:D3", "A1:A3", "B1:D3")] // Insert to left side - don't move
    [Arguments("A2:C4", "A1:C1", "A3:C5")] // Insert to top side - shift
    [Arguments("A2:C4", "A2:C2", "A3:C5")] // Insert to top edge - shift
    [Arguments("A2:C4", "A1", "B2:C4 A3:A5")] // Insert to top side - shift
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Insert to right side - don't move
    [Arguments("A1:C3", "A4:C5", "A1:C5")] // Insert to bottom edge - extend
    [Arguments("A1:C3", "A4", "A1:C3 A4")] // Insert to bottom side - extend
    [Arguments("A1:C3", "B4:E5", "A1:C3 B4:C5")] // Insert to bottom edge (inserted area is out of bounds of the area) - extend
    [Arguments("A1048576", "A1048576", "")] // Push out of sheet
    [Arguments("A1048575:A1048576", "A1048575", "A1048576")] // Partially push out of sheet
    [Arguments("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public void InsertAndShiftDown(string areaList, string insertedArea, string expected)
    {
        XLAreaList list = new(Area.Parse(areaList));
        XLAreaList result = list.InsertAndShiftDown(Area.Parse(insertedArea));

        ClassicAssert.AreEqual(expected, result.ToSpaceList());
    }

    [Test]
    public void InsertAndShiftDown_baseline_comparison()
    {
        // Compare the result of the method with the behavior of CFs Applied To field collected from Excel
        foreach (
            (XLAreaList original, Area insertArea, XLAreaList expectedResult) in GetBaselineData(
                "Other.ConditionalFormats.insert-and-shift-down-cf-baseline.txt"
            )
        )
        {
            XLAreaList result = original.InsertAndShiftDown(insertArea);
            ClassicAssert.AreEqual(expectedResult.ToSpaceList(), result.ToSpaceList());
        }
    }

    [Test]
    [Arguments("A1:C3", "A1", "A2:C3 B1:D1")]
    [Arguments("A1:C3", "B1", "A2:C3 A1:D1")]
    [Arguments("A1:C3", "C1", "A2:C3 A1:D1")]
    [Arguments("A1:C3", "A2", "A1:C1 A3:C3 B2:D2")]
    [Arguments("A1:C3", "B2", "A1:C1 A3:C3 A2:D2")]
    [Arguments("A1:C3", "C2", "A1:C1 A3:C3 A2:D2")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:D3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3:D3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:D3")]
    [Arguments("A1:C3", "A1:A3", "B1:D3")] // Insert to left edge - shift, don't extend
    [Arguments("A2:C4", "A1", "A2:C4")] // Insert to top side - don't move
    [Arguments("A1:C3", "D1:D3", "A1:D3")] // Insert to right edge - extend
    [Arguments("A1:C3", "D2:E10", "A1:C3 D2:E3")] // Insert to right edge (inserted area is out of bounds of the area) - extend
    [Arguments("A1:C3", "E1:E3", "A1:C3")] // Insert to right side  - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Insert to bottom side  - don't move
    [Arguments("XFD1", "XFD1", "")] // Push out of sheet
    [Arguments("XFC1:XFD1", "XFC1", "XFD1")] // Partially push out of sheet
    [Arguments("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public void InsertAndShiftRight(string areaList, string insertedArea, string expected)
    {
        XLAreaList list = new([Area.Parse(areaList)]);
        XLAreaList result = list.InsertAndShiftRight(Area.Parse(insertedArea));

        ClassicAssert.AreEqual(expected, result.ToSpaceList());
    }

    [Test]
    public void InsertAndShiftRight_baseline_comparison()
    {
        // Compare the result of the method with the behavior of CFs Applied To field collected from Excel
        foreach (
            (XLAreaList original, Area insertArea, XLAreaList expectedResult) in GetBaselineData(
                "Other.ConditionalFormats.insert-and-shift-right-cf-baseline.txt"
            )
        )
        {
            XLAreaList result = original.InsertAndShiftRight(insertArea);
            ClassicAssert.AreEqual(expectedResult.ToSpaceList(), result.ToSpaceList());
        }
    }

    [Test]
    [Arguments("A1:C3", "A1", "B1:C3 A1:A2")]
    [Arguments("A1:C3", "B1", "A1:A3 C1:C3 B1:B2")]
    [Arguments("A1:C3", "C1", "A1:B3 C1:C2")]
    [Arguments("A1:C3", "A2", "A1:C1 B2:C3 A2")]
    [Arguments("A1:C3", "B2", "A1:C1 A2:A3 C2:C3 B2")]
    [Arguments("A1:C3", "C2", "A1:C1 A2:B3 C2")]
    [Arguments("A1:C3", "A3", "A1:C2 B3:C3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 C3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3")]
    [Arguments("B1:D3", "A1:A3", "B1:D3")] // Delete on the left side - don't move
    [Arguments("A2:C4", "A1:C1", "A1:C3")] // Delete on top side - shift
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move
    [Arguments("A1:A3", "A1:D5", "")] // Delete completely
    [Arguments("A1:A1048576", "A1", "A1:A1048576")] // Columns are not changed
    public void DeleteAndShiftUp(string areaList, string deletedArea, string expected)
    {
        XLAreaList list = new([Area.Parse(areaList)]);
        XLAreaList result = list.DeleteAndShiftUp(Area.Parse(deletedArea));

        ClassicAssert.AreEqual(expected, result.ToSpaceList());
    }

    [Test]
    [Arguments("A1:C3", "A1", "A2:C3 A1:B1")]
    [Arguments("A1:C3", "B1", "A2:C3 A1 B1")]
    [Arguments("A1:C3", "C1", "A2:C3 A1:B1")]
    [Arguments("A1:C3", "A2", "A1:C1 A3:C3 A2:B2")]
    [Arguments("A1:C3", "B2", "A1:C1 A3:C3 A2 B2")]
    [Arguments("A1:C3", "C2", "A1:C1 A3:C3 A2:B2")]
    [Arguments("A1:C3", "A3", "A1:C2 A3:B3")]
    [Arguments("A1:C3", "B3", "A1:C2 A3 B3")]
    [Arguments("A1:C3", "C3", "A1:C2 A3:B3")]
    [Arguments("B1:D3", "A1:A3", "A1:C3")] // Delete on the left side - shift
    [Arguments("A2:C4", "A1", "A2:C4")] // Delete on top side - don't move
    [Arguments("A1:C3", "D1:D3", "A1:C3")] // Delete on right side - don't move
    [Arguments("A1:C3", "A4", "A1:C3")] // Delete on bottom side - don't move
    [Arguments("A1:A3", "A1:D5", "")] // Delete completely
    [Arguments("A1:XFD1", "A1", "A1:XFD1")] // Rows are not changed
    public void DeleteAndShiftLeft(string areaList, string deletedArea, string expected)
    {
        XLAreaList list = new([Area.Parse(areaList)]);
        XLAreaList result = list.DeleteAndShiftLeft(Area.Parse(deletedArea));

        ClassicAssert.AreEqual(expected, result.ToSpaceList());
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("A1:C3", "B2", true)]
    [Arguments("B2:C3", "A2", false)]
    [Arguments("A1:C2 B3:C3", "A3", false)]
    public void IntersectsWith_determines_intersection_with_any_area(
        string areaListText,
        string areaText,
        bool expected
    )
    {
        XLAreaList areaList = Parse(areaListText);
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(expected, areaList.IntersectsWith(area));
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1:C3", "B2", "A1:C3")]
    [Arguments("A1:C3", "B2:D4", "A1:C3")]
    [Arguments("A1 C1", "A1:C1", "A1 C1")]
    [Arguments("A1 C1", "B1", "")]
    [Arguments("A1 C1", "B1:D2", "C1")]
    public void IntersectingWith_returns_areas_intersecting_with_the_other_area(
        string areaListText,
        string areaText,
        string expected
    )
    {
        XLAreaList areaList = Parse(areaListText);
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(expected, areaList.IntersectingWith(area).ToSpaceList());
    }

    [Test]
    [Arguments("A1", "B1", "A1")]
    [Arguments("A1:E5", "C3:C4", "A1:E2 A5:E5 A3:B4 D3:E4")]
    [Arguments("B2:C5 B9 C4:D7", "C4:C5", "B2:C3 B4:B5 B9 C6:D7 D4:D5")]
    public void Excluding_returns_area_list_without_excluded(
        string areaListText,
        string excludedAreaText,
        string expected
    )
    {
        XLAreaList areaList = Parse(areaListText);
        Area excludedArea = Area.Parse(excludedAreaText);
        ClassicAssert.AreEqual(expected, areaList.Excluding(excludedArea).ToSpaceList());
    }

    [Test]
    [Arguments("A1", "A1", "A1", "A1")] // Copy from same point to the same point
    [Arguments("A1", "B5", "A1", "B5")] // Copy to different point
    [Arguments("B2", "D2", "A1:C3", "E3")] // The intersected area is not in corner and shifted doesn't start at the target point
    [Arguments("D3:G6", "A1", "E4:F5", "A1:B2")]
    [Arguments("B2", XLHelper.LastSheetAddress, "A1:C3", null)] // Copied area is out of sheet. Rare, but can happen.
    public void TryCopyAreaTo_return_list_of_intersecting_areas_shifted_to_target(
        string areaListText,
        string targetPointText,
        string areaToCopyText,
        string expected
    )
    {
        XLAreaList areaList = Parse(areaListText);
        Point targetPoint = Point.Parse(targetPointText);
        Area areaToCopy = Area.Parse(areaToCopyText);
        ClassicAssert.AreEqual(
            expected,
            areaList.TryCopyAreaTo(targetPoint, areaToCopy, out XLAreaList? result)
                ? result.ToSpaceList()
                : null
        );
    }

    private static XLAreaList Parse(string spaceList)
    {
        List<Area> list = [];
        foreach (string reference in spaceList.Split(' '))
        {
            list.Add(Area.Parse(reference));
        }

        return new XLAreaList(list);
    }

    private static IEnumerable<(XLAreaList, Area, XLAreaList)> GetBaselineData(string resourcePath)
    {
        using Stream stream = TestHelper.GetStreamFromResource(resourcePath);
        using StreamReader streamReader = new(stream);
        while (streamReader.ReadLine() is { } line)
        {
            string[] fields = line.Split(',');
            XLAreaList original = Parse(fields[0]);
            Area area = Area.Parse(fields[1]);
            XLAreaList expectedResult = Parse(fields[2]);
            yield return (original, area, expectedResult);
        }
    }
}

#nullable enable
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Coordinates;

public class AreaTests
{
    [Test]
    [Arguments("A1", 1, 1, 1, 1)]
    [Arguments("A1:Z100", 1, 1, 100, 26)]
    [Arguments("BD14:EG256", 14, 56, 256, 137)]
    [Arguments("A1:XFD1048576", 1, 1, 1048576, 16384)]
    [Arguments("XFD1048576", 1048576, 16384, 1048576, 16384)]
    [Arguments("XFD1048576:XFD1048576", 1048576, 16384, 1048576, 16384)]
    public void ParseCellRefsAccordingToGrammar(
        string refText,
        int firstRow,
        int firstCol,
        int lastRow,
        int lastCol
    )
    {
        Area reference = Area.Parse(refText);
        ClassicAssert.AreEqual(firstRow, reference.FirstPoint.Row);
        ClassicAssert.AreEqual(firstCol, reference.FirstPoint.Column);
        ClassicAssert.AreEqual(lastRow, reference.LastPoint.Row);
        ClassicAssert.AreEqual(lastCol, reference.LastPoint.Column);
    }

    [Test]
    [Arguments("")]
    [Arguments("A1:")]
    [Arguments(":A1")]
    [Arguments("A1: A1")]
    [Arguments(" A1:A1")]
    [Arguments("A1:A1 ")]
    [Arguments("B1:A1")]
    [Arguments("A2:A1")]
    public void InvalidInputsAreNotParsed(string invalidRef) =>
        ClassicAssert.Throws<FormatException>(() => Area.Parse(invalidRef));

    [Test]
    [Arguments("A1:A1", "A1")]
    [Arguments("DO974:LAR2487", "DO974:LAR2487")]
    [Arguments("XFD1048576:XFD1048576", "XFD1048576")]
    [Arguments("XFD1048575:XFD1048576", "XFD1048575:XFD1048576")]
    public void CanFormatToString(string cellRef, string expected)
    {
        Area r = Area.Parse(cellRef);
        ClassicAssert.AreEqual(expected, r.ToString());
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1", "B3", "A1:B3")]
    [Arguments("C2", "B3", "B2:C3")]
    [Arguments("I6:J9", "L7", "I6:L9")]
    [Arguments("B2:B4", "A3:C3", "A2:C4")]
    [Arguments("B2:C3", "E5:F6", "B2:F6")]
    public void RangeOperation(string leftOperand, string rightOperand, string expectedRange)
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);
        Area expected = Area.Parse(expectedRange);

        ClassicAssert.AreEqual(expected, left.Range(right));
    }

    [Test]
    [Arguments("A1", "A1", "A1")]
    [Arguments("A1", "A2", null)]
    [Arguments("B1:B3", "A2:C2", "B2")]
    [Arguments("A1:A3", "B2:C2", null)]
    [Arguments("A1:D6", "B2:C3", "B2:C3")]
    [Arguments("A1:C6", "B4:E10", "B4:C6")]
    public void IntersectOperation(string leftOperand, string rightOperand, string? expectedRange)
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);
        Area? expected = expectedRange is null ? (Area?)null : Area.Parse(expectedRange);

        ClassicAssert.AreEqual(expected, left.Intersect(right));
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("A1", "A2", false)]
    [Arguments("B1:B3", "A2:C2", true)]
    [Arguments("A1:A3", "B2:C2", false)]
    [Arguments("A1:D6", "B2:C3", true)]
    [Arguments("A1:C6", "B4:E10", true)]
    public void IntersectsChecksWhetherTheRangeHasIntersectionWithAnother(
        string leftOperand,
        string rightOperand,
        bool expected
    )
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);

        ClassicAssert.AreEqual(expected, left.Intersects(right));
    }

    [Test]
    [Arguments("A1", "A1", true)]
    [Arguments("B1:C3", "B1:C3", true)]
    [Arguments("A1:D4", "B2:C3", true)]
    [Arguments("B3:C3", "B2:C3", false)]
    [Arguments("A2:C2", "B2:C3", false)]
    public void OverlapsChecksWhetherLeftFullyOverlapsRight(
        string leftOperand,
        string rightOperand,
        bool expected
    )
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);

        ClassicAssert.AreEqual(expected, left.Overlaps(right));
    }

    [Test]
    [Arguments("C4:F8", "C1:F3", "C4:F8")] // Inserted area is fully above
    [Arguments("C4:F8", "A9:G12", "C4:F8")] // Inserted area is fully below
    [Arguments("C4:F8", "G1:H5", "C4:F8")] // Inserted are is fully to the right
    [Arguments("C4:F8", "C1:D11", "E4:H8")] // Inserted area at the left column of the area
    [Arguments("C4:F8", "A1:B8", "E4:H8")] // Inserted area is fully to the left
    [Arguments("C4:F8", "D4:E8", "C4:H8")] // Inserted into the area
    [Arguments("C4:F8", "D2:I8", "C4:L8")] // Inside the area, overlapping = extend
    [Arguments("C4:F8", "F4:F8", "C4:G8")] // Last column of the area, overlapping = extend
    [Arguments("XFD1", "XFB1", null)] // Completely pushed out of the range
    [Arguments("XFA1:XFD1", "XEZ1:XFA1", "XFC1:XFD1")] // Partially pushed out of the range
    [Arguments("XFA1:XFD1", "XFB1:XFC1", "XFA1:XFD1")] // Extend below last row
    public void TryInsertAreaAndShiftRightWithoutPartialCover(
        string original,
        string inserted,
        string? repositioned
    )
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);
        Area? repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        bool success = originalArea.TryInsertAreaAndShiftRight(insertedArea, out Area? result);

        ClassicAssert.True(success);
        ClassicAssert.AreEqual(repositionedArea, result);
    }

    [Test]
    [Arguments("C4:F8", "B3:B4")] // Partially above
    [Arguments("C4:F8", "B5:C7")] // In the middle
    [Arguments("C4:F8", "A5:B9")] // Partially below
    public void TryInsertAreaAndShiftRightWithPartialCover(string original, string inserted)
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);

        ClassicAssert.False(originalArea.TryInsertAreaAndShiftRight(insertedArea, out _));
    }

    [Test]
    [Arguments("D6:G10", "A1:C15", "D6:G10")] // Inserted are is fully to the left
    [Arguments("D6:G10", "H1:K15", "D6:G10")] // Inserted are is fully to the right
    [Arguments("D6:G10", "A11:K15", "D6:G10")] // Inserted are is fully below
    [Arguments("D6:G10", "D6:G11", "D12:G16")] // Inserted area at the top row of the area
    [Arguments("D6:G10", "C4:H7", "D10:G14")] // Inserted above the area
    [Arguments("D6:G10", "D7:G9", "D6:G13")] // Inserted into the area
    [Arguments("D6:G10", "A7:H9", "D6:G13")] // Inside the area, overlapping = extend
    [Arguments("D6:G10", "D10:G11", "D6:G12")] // Last row of the area, overlapping = extend
    [Arguments("A1048576", "A1048575", null)] // Completely pushed out of the range
    [Arguments("A1048574:A1048576", "A1048570:A1048571", "A1048576")] // Partially pushed out of the range
    [Arguments("A1048570:A1048572", "A1048571:A1048576", "A1048570:A1048576")] // Extend below last row
    public void TryInsertAreaAndShiftDownWithoutPartialCover(
        string original,
        string inserted,
        string? repositioned
    )
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);
        Area? repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        bool success = originalArea.TryInsertAreaAndShiftDown(insertedArea, out Area? result);

        ClassicAssert.True(success);
        ClassicAssert.AreEqual(repositionedArea, result);
    }

    [Test]
    [Arguments("D6:G10", "A6:E6")] // Left
    [Arguments("D6:G10", "D5:D5")] // Above
    [Arguments("D6:G10", "E7:H15")] // Right
    public void TryInsertAreaAndShiftDownWithPartialCover(string original, string inserted)
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);

        ClassicAssert.False(originalArea.TryInsertAreaAndShiftDown(insertedArea, out _));
    }

    [Test]
    [Arguments("E4:G4", "B3:C5", "C4:E4")] // Deleted area fully to the left with overlapping width
    [Arguments("E4:G4", "A2:D5", "A4:C4")] // The deleted are ends exactly at the column to the left of the area
    [Arguments("E4:G4", "F1:F7", "E4:F4")] // The deleted is fully within the area, but not at left/right column
    [Arguments("E4:G4", "E4:G4", null)] // Delete are exactly covers the area
    [Arguments("E4:G4", "A1:Z9", null)] // Delete fully covers the area
    [Arguments("E4:G4", "H1:K10", "E4:G4")] // The deleted is fully to the right of the area.
    [Arguments("E4:G4", "G3:H5", "E4:F4")] // The deleted partially intersects the area and is to the right.
    [Arguments("D4:E4", "A5:F9", "D4:E4")] // Deleted area is fully downward
    [Arguments("D4:E4", "A1:F3", "D4:E4")] // Deleted area is fully upwards
    [Arguments("D4:E4", "A5:F10", "D4:E4")] // Partial deletion is below -> not affected
    [Arguments("D4:F8", "D4:F6", "D7:F8")] // Delete top slice
    [Arguments("D4:F8", "B1:H6", "D7:F8")] // Delete top slice
    [Arguments("D4:F8", "D6:F8", "D4:F5")] // Delete bottom slice
    [Arguments("D4:F8", "B6:I15", "D4:F5")] // Delete bottom slice
    public void TryDeleteAreaAndShiftLeftWithoutPartialCover(
        string original,
        string deleted,
        string? repositioned
    )
    {
        Area originalArea = Area.Parse(original);
        Area deletedArea = Area.Parse(deleted);
        Area? repositionedArea = repositioned is not null ? Area.Parse(repositioned) : (Area?)null;

        bool success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out Area? result);

        ClassicAssert.True(success);
        ClassicAssert.AreEqual(repositionedArea, result);
    }

    [Test]
    [Arguments("D4:E8", "A1:B5")] // Partial left
    [Arguments("D4:E8", "D6:E7")] // Partial inside
    [Arguments("D4:E8", "C4:D6")] // Partial left and inside
    public void TryDeleteAreaAndShiftLeftWithPartialCover(string original, string deleted)
    {
        Area originalArea = Area.Parse(original);
        Area deletedArea = Area.Parse(deleted);
        bool success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out Area? result);

        ClassicAssert.False(success);
        ClassicAssert.Null(result);
    }

    [Test]
    [Arguments("B5:B8", "A1:C3", "B2:B5")] // Deleted area fully above (with a row space) with overlapping width
    [Arguments("B5:B8", "A2:C4", "B2:B5")] // The deleted are ends exactly at the row above the area
    [Arguments("B5:B8", "A6:C7", "B5:B6")] // The deleted is fully within the area, but not at top/bottom row
    [Arguments("B5:B8", "A5:C8", null)] // Delete are exactly covers the area
    [Arguments("B5:B8", "A4:C9", null)] // Delete fully covers the area
    [Arguments("B5:B8", "A9:C10", "B5:B8")] // The deleted is fully below the area.
    [Arguments("B5:B8", "A6:C10", "B5:B5")] // The deleted partially intersects the area and is below.
    [Arguments("B5:B8", "A1:A10", "B5:B8")] // Deleted area is fully on the left
    [Arguments("B5:B8", "C1:C10", "B5:B8")] // Deleted area is fully on the right
    [Arguments("B5:D8", "B9:C10", "B5:D8")] // Partial deletion is below -> not affected
    [Arguments("D4:H8", "D4:F8", "G4:H8")] // Delete left slice
    [Arguments("D4:H8", "C1:F9", "G4:H8")] // Delete left slice
    [Arguments("D4:H8", "G4:H8", "D4:F8")] // Delete right slice
    [Arguments("D4:H8", "G1:I9", "D4:F8")] // Delete right slice
    public void TryDeleteAreaAndShiftUpWithoutPartialCover(
        string leftOperand,
        string deleted,
        string? expected
    )
    {
        Area originalArea = Area.Parse(leftOperand);
        Area deletedArea = Area.Parse(deleted);
        Area? expectedResult = expected is not null ? Area.Parse(expected) : (Area?)null;

        bool success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out Area? result);

        ClassicAssert.True(success);
        ClassicAssert.AreEqual(expectedResult, result);
    }

    [Test]
    [Arguments("B5:D8", "A1:B3")] // Partial above
    [Arguments("B5:D8", "C6:D8")] // Partial inside
    [Arguments("B5:D8", "B1:B6")] // Partial above and inside
    public void TryDeleteAreaAndShiftUpWithPartialCover(string leftOperand, string deleted)
    {
        Area originalArea = Area.Parse(leftOperand);
        Area deletedArea = Area.Parse(deleted);
        bool success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out Area? result);

        ClassicAssert.False(success);
        ClassicAssert.Null(result);
    }

    [Test]
    [Arguments("B2:D4", "B2", "B2", "B3:D4 C2:D2")]
    [Arguments("B2:D4", "A1:B2", "B2", "B3:D4 C2:D2")]
    [Arguments("B2:D4", "C2", "C2", "B3:D4 B2 D2")]
    [Arguments("B2:D4", "C1:C2", "C2", "B3:D4 B2 D2")]
    [Arguments("B2:D4", "D2", "D2", "B3:D4 B2:C2")]
    [Arguments("B2:D4", "D1:E2", "D2", "B3:D4 B2:C2")]
    [Arguments("B2:D4", "B3", "B3", "B2:D2 B4:D4 C3:D3")]
    [Arguments("B2:D4", "A3:B3", "B3", "B2:D2 B4:D4 C3:D3")]
    [Arguments("B2:D4", "C3", "C3", "B2:D2 B4:D4 B3 D3")]
    [Arguments("B2:D4", "D3", "D3", "B2:D2 B4:D4 B3:C3")]
    [Arguments("B2:D4", "D3:E3", "D3", "B2:D2 B4:D4 B3:C3")]
    [Arguments("B2:D4", "B4", "B4", "B2:D3 C4:D4")]
    [Arguments("B2:D4", "A4:B5", "B4", "B2:D3 C4:D4")]
    [Arguments("B2:D4", "C4", "C4", "B2:D3 B4 D4")]
    [Arguments("B2:D4", "C4:C5", "C4", "B2:D3 B4 D4")]
    [Arguments("B2:D4", "D4", "D4", "B2:D3 B4:C4")]
    [Arguments("B2:D4", "D4:E5", "D4", "B2:D3 B4:C4")]
    [Arguments("B2:D4", "B3:D3", "B3:D3", "B2:D2 B4:D4")]
    [Arguments("B2:D4", "A3:E3", "B3:D3", "B2:D2 B4:D4")]
    [Arguments("B2:D4", "C2:C4", "C2:C4", "B2:B4 D2:D4")]
    [Arguments("B2:D4", "C1:C5", "C2:C4", "B2:B4 D2:D4")]
    public void ExcludeSplitsOriginalAreaWhenExcludedAreaIntersects(
        string originalAreaText,
        string excludingAreaText,
        string excludedAreaText,
        string expected
    )
    {
        Area originalArea = Area.Parse(originalAreaText);
        Area excludedArea = Area.Parse(excludedAreaText);
        Area excludingArea = Area.Parse(excludingAreaText);
        List<Area> list = [];
        ClassicAssert.AreEqual(excludedArea, originalArea.Exclude(excludingArea, list));
        ClassicAssert.AreEqual(expected, list.ToSpaceList());
    }

    [Test]
    [Arguments("B2:C3", "A1")]
    [Arguments("B2:C3", "D1:G20")]
    [Arguments("A1", "A2:C5")]
    public void ExcludeKeepsOriginalAreaWhenExcludedAreaDoesntIntersects(
        string originalAreaText,
        string excludedAreaText
    )
    {
        Area originalArea = Area.Parse(originalAreaText);
        Area excludedArea = Area.Parse(excludedAreaText);
        List<Area> list = [];
        ClassicAssert.IsNull(originalArea.Exclude(excludedArea, list));
        ClassicAssert.AreEqual(originalAreaText, list.ToSpaceList());
    }

    [Test]
    [Arguments("A1", 0, 2, "C1")]
    [Arguments("A1", 1, 0, "A2")]
    [Arguments("A1", 0, -1, null)]
    [Arguments("A1", -1, 0, null)]
    [Arguments(XLHelper.LastSheetAddress, 0, 1, null)]
    [Arguments(XLHelper.LastSheetAddress, 1, 0, null)]
    [Arguments("B2:D3", 2, 3, "E4:G5")]
    [Arguments("B2:D3", -2, -3, "A1")]
    [Arguments("XFA1048574:XFC1048575", 2, 3, "XFD1048576")]
    public void ShiftAndClipShiftsAndClipsArea(
        string areaText,
        int rowShift,
        int columnShift,
        string? expected
    )
    {
        Area area = Area.Parse(areaText);
        Area? shifted = area.ShiftAndClip(rowShift, columnShift);
        ClassicAssert.AreEqual(expected, shifted is not null ? shifted.ToString() : null);
    }

    [Test]
    [Arguments("A1:A2", 2, true, "A1", "A2")]
    [Arguments("A1", 2, true, "A1", null)]
    [Arguments("A1", 1, false, null, "A1")]
    public void SplitAboveSplitsAreaIntoAreaAboveAndBelowTheRow(
        string areaText,
        int row,
        bool isAbove,
        string? expectedAbove,
        string? expectedBelow
    )
    {
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(isAbove, area.SplitAbove(row, out Area? above, out Area? below));
        ClassicAssert.AreEqual(expectedAbove, above?.ToString());
        ClassicAssert.AreEqual(expectedBelow, below?.ToString());
    }

    [Test]
    [Arguments("A1:A2", 1, true, "A2", "A1")]
    [Arguments("A1", 2, false, null, "A1")]
    [Arguments("A2", 1, true, "A2", null)]
    public void SplitBelowSplitsAreaIntoAreaBelowAndAboveTheRow(
        string areaText,
        int row,
        bool isBelow,
        string? expectedBelow,
        string? expectedAbove
    )
    {
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(isBelow, area.SplitBelow(row, out Area? below, out Area? above));
        ClassicAssert.AreEqual(expectedBelow, below?.ToString());
        ClassicAssert.AreEqual(expectedAbove, above?.ToString());
    }

    [Test]
    [Arguments("A1:B1", 2, true, "A1", "B1")]
    [Arguments("A1", 2, true, "A1", null)]
    [Arguments("A1:C3", 1, false, null, "A1:C3")]
    public void SplitBeforeSplitsAreaIntoAreaToLeftAndRightOfColumn(
        string areaText,
        int column,
        bool isToLeft,
        string? expectedLeft,
        string? expectedRight
    )
    {
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(isToLeft, area.SplitBefore(column, out Area? left, out Area? right));
        ClassicAssert.AreEqual(expectedLeft, left?.ToString());
        ClassicAssert.AreEqual(expectedRight, right?.ToString());
    }

    [Test]
    [Arguments("A1:B1", 1, true, "B1", "A1")]
    [Arguments("B2:C3", 1, true, "B2:C3", null)]
    [Arguments("A1", 1, false, null, "A1")]
    public void SplitAfterSplitsAreaIntoAreaToRightAndLeftOfColumn(
        string areaText,
        int column,
        bool isToRight,
        string? expectedRight,
        string? expectedLeft
    )
    {
        Area area = Area.Parse(areaText);
        ClassicAssert.AreEqual(isToRight, area.SplitAfter(column, out Area? right, out Area? left));
        ClassicAssert.AreEqual(expectedRight, right?.ToString());
        ClassicAssert.AreEqual(expectedLeft, left?.ToString());
    }

    [Test]
    public async Task Single_cell_areas_hash_codes_have_few_collisions()
    {
        // Areas are used in many dictionaries, such as formula dependency tree. Make sure the hash
        // function doesn't produce too many collision. The hash function originally always
        // produced hash code 0 for all one-cell areas, which caused a terrible set/dictionary
        // performance.
        var hashes = new HashSet<int>();
        for (var row = 1; row <= 100; ++row)
        {
            for (var column = 1; column <= 100; ++column)
                hashes.Add(new Area(new Point(row, column)).GetHashCode());
        }

        // Some collisions are inevitable, but the vast majority of areas must be distinguishable.
        await Assert.That(hashes.Count).IsGreaterThan(9900);
    }
}

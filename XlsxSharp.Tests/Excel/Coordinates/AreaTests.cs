#nullable enable
using System;
using System.Collections.Generic;
using NUnit.Framework;
using XlsxSharp.Excel;
using XlsxSharp.Tests.Utils;

namespace XlsxSharp.Tests.Excel.Coordinates;

[TestFixture]
public class AreaTests
{
    [TestCase("A1", 1, 1, 1, 1)]
    [TestCase("A1:Z100", 1, 1, 100, 26)]
    [TestCase("BD14:EG256", 14, 56, 256, 137)]
    [TestCase("A1:XFD1048576", 1, 1, 1048576, 16384)]
    [TestCase("XFD1048576", 1048576, 16384, 1048576, 16384)]
    [TestCase("XFD1048576:XFD1048576", 1048576, 16384, 1048576, 16384)]
    public void ParseCellRefsAccordingToGrammar(
        string refText,
        int firstRow,
        int firstCol,
        int lastRow,
        int lastCol
    )
    {
        Area reference = Area.Parse(refText);
        Assert.AreEqual(firstRow, reference.FirstPoint.Row);
        Assert.AreEqual(firstCol, reference.FirstPoint.Column);
        Assert.AreEqual(lastRow, reference.LastPoint.Row);
        Assert.AreEqual(lastCol, reference.LastPoint.Column);
    }

    [TestCase("")]
    [TestCase("A1:")]
    [TestCase(":A1")]
    [TestCase("A1: A1")]
    [TestCase(" A1:A1")]
    [TestCase("A1:A1 ")]
    [TestCase("B1:A1")]
    [TestCase("A2:A1")]
    public void InvalidInputsAreNotParsed(string invalidRef) =>
        Assert.Throws<FormatException>(() => Area.Parse(invalidRef));

    [TestCase("A1:A1", "A1")]
    [TestCase("DO974:LAR2487", "DO974:LAR2487")]
    [TestCase("XFD1048576:XFD1048576", "XFD1048576")]
    [TestCase("XFD1048575:XFD1048576", "XFD1048575:XFD1048576")]
    public void CanFormatToString(string cellRef, string expected)
    {
        Area r = Area.Parse(cellRef);
        Assert.AreEqual(expected, r.ToString());
    }

    [TestCase("A1", "A1", "A1")]
    [TestCase("A1", "B3", "A1:B3")]
    [TestCase("C2", "B3", "B2:C3")]
    [TestCase("I6:J9", "L7", "I6:L9")]
    [TestCase("B2:B4", "A3:C3", "A2:C4")]
    [TestCase("B2:C3", "E5:F6", "B2:F6")]
    public void RangeOperation(string leftOperand, string rightOperand, string expectedRange)
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);
        Area expected = Area.Parse(expectedRange);

        Assert.AreEqual(expected, left.Range(right));
    }

    [TestCase("A1", "A1", "A1")]
    [TestCase("A1", "A2", null)]
    [TestCase("B1:B3", "A2:C2", "B2")]
    [TestCase("A1:A3", "B2:C2", null)]
    [TestCase("A1:D6", "B2:C3", "B2:C3")]
    [TestCase("A1:C6", "B4:E10", "B4:C6")]
    public void IntersectOperation(string leftOperand, string rightOperand, string? expectedRange)
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);
        Area? expected = expectedRange is null ? (Area?)null : Area.Parse(expectedRange);

        Assert.AreEqual(expected, left.Intersect(right));
    }

    [TestCase("A1", "A1", true)]
    [TestCase("A1", "A2", false)]
    [TestCase("B1:B3", "A2:C2", true)]
    [TestCase("A1:A3", "B2:C2", false)]
    [TestCase("A1:D6", "B2:C3", true)]
    [TestCase("A1:C6", "B4:E10", true)]
    public void IntersectsChecksWhetherTheRangeHasIntersectionWithAnother(
        string leftOperand,
        string rightOperand,
        bool expected
    )
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);

        Assert.AreEqual(expected, left.Intersects(right));
    }

    [TestCase("A1", "A1", true)]
    [TestCase("B1:C3", "B1:C3", true)]
    [TestCase("A1:D4", "B2:C3", true)]
    [TestCase("B3:C3", "B2:C3", false)]
    [TestCase("A2:C2", "B2:C3", false)]
    public void OverlapsChecksWhetherLeftFullyOverlapsRight(
        string leftOperand,
        string rightOperand,
        bool expected
    )
    {
        Area left = Area.Parse(leftOperand);
        Area right = Area.Parse(rightOperand);

        Assert.AreEqual(expected, left.Overlaps(right));
    }

    [TestCase("C4:F8", "C1:F3", "C4:F8")] // Inserted area is fully above
    [TestCase("C4:F8", "A9:G12", "C4:F8")] // Inserted area is fully below
    [TestCase("C4:F8", "G1:H5", "C4:F8")] // Inserted are is fully to the right
    [TestCase("C4:F8", "C1:D11", "E4:H8")] // Inserted area at the left column of the area
    [TestCase("C4:F8", "A1:B8", "E4:H8")] // Inserted area is fully to the left
    [TestCase("C4:F8", "D4:E8", "C4:H8")] // Inserted into the area
    [TestCase("C4:F8", "D2:I8", "C4:L8")] // Inside the area, overlapping = extend
    [TestCase("C4:F8", "F4:F8", "C4:G8")] // Last column of the area, overlapping = extend
    [TestCase("XFD1", "XFB1", null)] // Completely pushed out of the range
    [TestCase("XFA1:XFD1", "XEZ1:XFA1", "XFC1:XFD1")] // Partially pushed out of the range
    [TestCase("XFA1:XFD1", "XFB1:XFC1", "XFA1:XFD1")] // Extend below last row
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

        Assert.True(success);
        Assert.AreEqual(repositionedArea, result);
    }

    [TestCase("C4:F8", "B3:B4")] // Partially above
    [TestCase("C4:F8", "B5:C7")] // In the middle
    [TestCase("C4:F8", "A5:B9")] // Partially below
    public void TryInsertAreaAndShiftRightWithPartialCover(string original, string inserted)
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);

        Assert.False(originalArea.TryInsertAreaAndShiftRight(insertedArea, out _));
    }

    [TestCase("D6:G10", "A1:C15", "D6:G10")] // Inserted are is fully to the left
    [TestCase("D6:G10", "H1:K15", "D6:G10")] // Inserted are is fully to the right
    [TestCase("D6:G10", "A11:K15", "D6:G10")] // Inserted are is fully below
    [TestCase("D6:G10", "D6:G11", "D12:G16")] // Inserted area at the top row of the area
    [TestCase("D6:G10", "C4:H7", "D10:G14")] // Inserted above the area
    [TestCase("D6:G10", "D7:G9", "D6:G13")] // Inserted into the area
    [TestCase("D6:G10", "A7:H9", "D6:G13")] // Inside the area, overlapping = extend
    [TestCase("D6:G10", "D10:G11", "D6:G12")] // Last row of the area, overlapping = extend
    [TestCase("A1048576", "A1048575", null)] // Completely pushed out of the range
    [TestCase("A1048574:A1048576", "A1048570:A1048571", "A1048576")] // Partially pushed out of the range
    [TestCase("A1048570:A1048572", "A1048571:A1048576", "A1048570:A1048576")] // Extend below last row
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

        Assert.True(success);
        Assert.AreEqual(repositionedArea, result);
    }

    [TestCase("D6:G10", "A6:E6")] // Left
    [TestCase("D6:G10", "D5:D5")] // Above
    [TestCase("D6:G10", "E7:H15")] // Right
    public void TryInsertAreaAndShiftDownWithPartialCover(string original, string inserted)
    {
        Area originalArea = Area.Parse(original);
        Area insertedArea = Area.Parse(inserted);

        Assert.False(originalArea.TryInsertAreaAndShiftDown(insertedArea, out _));
    }

    [TestCase("E4:G4", "B3:C5", "C4:E4")] // Deleted area fully to the left with overlapping width
    [TestCase("E4:G4", "A2:D5", "A4:C4")] // The deleted are ends exactly at the column to the left of the area
    [TestCase("E4:G4", "F1:F7", "E4:F4")] // The deleted is fully within the area, but not at left/right column
    [TestCase("E4:G4", "E4:G4", null)] // Delete are exactly covers the area
    [TestCase("E4:G4", "A1:Z9", null)] // Delete fully covers the area
    [TestCase("E4:G4", "H1:K10", "E4:G4")] // The deleted is fully to the right of the area.
    [TestCase("E4:G4", "G3:H5", "E4:F4")] // The deleted partially intersects the area and is to the right.
    [TestCase("D4:E4", "A5:F9", "D4:E4")] // Deleted area is fully downward
    [TestCase("D4:E4", "A1:F3", "D4:E4")] // Deleted area is fully upwards
    [TestCase("D4:E4", "A5:F10", "D4:E4")] // Partial deletion is below -> not affected
    [TestCase("D4:F8", "D4:F6", "D7:F8")] // Delete top slice
    [TestCase("D4:F8", "B1:H6", "D7:F8")] // Delete top slice
    [TestCase("D4:F8", "D6:F8", "D4:F5")] // Delete bottom slice
    [TestCase("D4:F8", "B6:I15", "D4:F5")] // Delete bottom slice
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

        Assert.True(success);
        Assert.AreEqual(repositionedArea, result);
    }

    [TestCase("D4:E8", "A1:B5")] // Partial left
    [TestCase("D4:E8", "D6:E7")] // Partial inside
    [TestCase("D4:E8", "C4:D6")] // Partial left and inside
    public void TryDeleteAreaAndShiftLeftWithPartialCover(string original, string deleted)
    {
        Area originalArea = Area.Parse(original);
        Area deletedArea = Area.Parse(deleted);
        bool success = originalArea.TryDeleteAreaAndShiftLeft(deletedArea, out Area? result);

        Assert.False(success);
        Assert.Null(result);
    }

    [TestCase("B5:B8", "A1:C3", "B2:B5")] // Deleted area fully above (with a row space) with overlapping width
    [TestCase("B5:B8", "A2:C4", "B2:B5")] // The deleted are ends exactly at the row above the area
    [TestCase("B5:B8", "A6:C7", "B5:B6")] // The deleted is fully within the area, but not at top/bottom row
    [TestCase("B5:B8", "A5:C8", null)] // Delete are exactly covers the area
    [TestCase("B5:B8", "A4:C9", null)] // Delete fully covers the area
    [TestCase("B5:B8", "A9:C10", "B5:B8")] // The deleted is fully below the area.
    [TestCase("B5:B8", "A6:C10", "B5:B5")] // The deleted partially intersects the area and is below.
    [TestCase("B5:B8", "A1:A10", "B5:B8")] // Deleted area is fully on the left
    [TestCase("B5:B8", "C1:C10", "B5:B8")] // Deleted area is fully on the right
    [TestCase("B5:D8", "B9:C10", "B5:D8")] // Partial deletion is below -> not affected
    [TestCase("D4:H8", "D4:F8", "G4:H8")] // Delete left slice
    [TestCase("D4:H8", "C1:F9", "G4:H8")] // Delete left slice
    [TestCase("D4:H8", "G4:H8", "D4:F8")] // Delete right slice
    [TestCase("D4:H8", "G1:I9", "D4:F8")] // Delete right slice
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

        Assert.True(success);
        Assert.AreEqual(expectedResult, result);
    }

    [TestCase("B5:D8", "A1:B3")] // Partial above
    [TestCase("B5:D8", "C6:D8")] // Partial inside
    [TestCase("B5:D8", "B1:B6")] // Partial above and inside
    public void TryDeleteAreaAndShiftUpWithPartialCover(string leftOperand, string deleted)
    {
        Area originalArea = Area.Parse(leftOperand);
        Area deletedArea = Area.Parse(deleted);
        bool success = originalArea.TryDeleteAreaAndShiftUp(deletedArea, out Area? result);

        Assert.False(success);
        Assert.Null(result);
    }

    [TestCase("B2:D4", "B2", "B2", ExpectedResult = "B3:D4 C2:D2")]
    [TestCase("B2:D4", "A1:B2", "B2", ExpectedResult = "B3:D4 C2:D2")]
    [TestCase("B2:D4", "C2", "C2", ExpectedResult = "B3:D4 B2 D2")]
    [TestCase("B2:D4", "C1:C2", "C2", ExpectedResult = "B3:D4 B2 D2")]
    [TestCase("B2:D4", "D2", "D2", ExpectedResult = "B3:D4 B2:C2")]
    [TestCase("B2:D4", "D1:E2", "D2", ExpectedResult = "B3:D4 B2:C2")]
    [TestCase("B2:D4", "B3", "B3", ExpectedResult = "B2:D2 B4:D4 C3:D3")]
    [TestCase("B2:D4", "A3:B3", "B3", ExpectedResult = "B2:D2 B4:D4 C3:D3")]
    [TestCase("B2:D4", "C3", "C3", ExpectedResult = "B2:D2 B4:D4 B3 D3")]
    [TestCase("B2:D4", "D3", "D3", ExpectedResult = "B2:D2 B4:D4 B3:C3")]
    [TestCase("B2:D4", "D3:E3", "D3", ExpectedResult = "B2:D2 B4:D4 B3:C3")]
    [TestCase("B2:D4", "B4", "B4", ExpectedResult = "B2:D3 C4:D4")]
    [TestCase("B2:D4", "A4:B5", "B4", ExpectedResult = "B2:D3 C4:D4")]
    [TestCase("B2:D4", "C4", "C4", ExpectedResult = "B2:D3 B4 D4")]
    [TestCase("B2:D4", "C4:C5", "C4", ExpectedResult = "B2:D3 B4 D4")]
    [TestCase("B2:D4", "D4", "D4", ExpectedResult = "B2:D3 B4:C4")]
    [TestCase("B2:D4", "D4:E5", "D4", ExpectedResult = "B2:D3 B4:C4")]
    [TestCase("B2:D4", "B3:D3", "B3:D3", ExpectedResult = "B2:D2 B4:D4")]
    [TestCase("B2:D4", "A3:E3", "B3:D3", ExpectedResult = "B2:D2 B4:D4")]
    [TestCase("B2:D4", "C2:C4", "C2:C4", ExpectedResult = "B2:B4 D2:D4")]
    [TestCase("B2:D4", "C1:C5", "C2:C4", ExpectedResult = "B2:B4 D2:D4")]
    public string ExcludeSplitsOriginalAreaWhenExcludedAreaIntersects(
        string originalAreaText,
        string excludingAreaText,
        string excludedAreaText
    )
    {
        Area originalArea = Area.Parse(originalAreaText);
        Area excludedArea = Area.Parse(excludedAreaText);
        Area excludingArea = Area.Parse(excludingAreaText);
        List<Area> list = [];
        Assert.AreEqual(excludedArea, originalArea.Exclude(excludingArea, list));
        return list.ToSpaceList();
    }

    [TestCase("B2:C3", "A1")]
    [TestCase("B2:C3", "D1:G20")]
    [TestCase("A1", "A2:C5")]
    public void ExcludeKeepsOriginalAreaWhenExcludedAreaDoesntIntersects(
        string originalAreaText,
        string excludedAreaText
    )
    {
        Area originalArea = Area.Parse(originalAreaText);
        Area excludedArea = Area.Parse(excludedAreaText);
        List<Area> list = [];
        Assert.IsNull(originalArea.Exclude(excludedArea, list));
        Assert.AreEqual(originalAreaText, list.ToSpaceList());
    }

    [TestCase("A1", 0, 2, ExpectedResult = "C1")]
    [TestCase("A1", 1, 0, ExpectedResult = "A2")]
    [TestCase("A1", 0, -1, ExpectedResult = null)]
    [TestCase("A1", -1, 0, ExpectedResult = null)]
    [TestCase(XLHelper.LastSheetAddress, 0, 1, ExpectedResult = null)]
    [TestCase(XLHelper.LastSheetAddress, 1, 0, ExpectedResult = null)]
    [TestCase("B2:D3", 2, 3, ExpectedResult = "E4:G5")]
    [TestCase("B2:D3", -2, -3, ExpectedResult = "A1")]
    [TestCase("XFA1048574:XFC1048575", 2, 3, ExpectedResult = "XFD1048576")]
    public string? ShiftAndClipShiftsAndClipsArea(string areaText, int rowShift, int columnShift)
    {
        Area area = Area.Parse(areaText);
        Area? shifted = area.ShiftAndClip(rowShift, columnShift);
        return shifted is not null ? shifted.ToString() : null;
    }

    [TestCase("A1:A2", 2, true, "A1", "A2")]
    [TestCase("A1", 2, true, "A1", null)]
    [TestCase("A1", 1, false, null, "A1")]
    public void SplitAboveSplitsAreaIntoAreaAboveAndBelowTheRow(
        string areaText,
        int row,
        bool isAbove,
        string? expectedAbove,
        string? expectedBelow
    )
    {
        Area area = Area.Parse(areaText);
        Assert.AreEqual(isAbove, area.SplitAbove(row, out Area? above, out Area? below));
        Assert.AreEqual(expectedAbove, above?.ToString());
        Assert.AreEqual(expectedBelow, below?.ToString());
    }

    [TestCase("A1:A2", 1, true, "A2", "A1")]
    [TestCase("A1", 2, false, null, "A1")]
    [TestCase("A2", 1, true, "A2", null)]
    public void SplitBelowSplitsAreaIntoAreaBelowAndAboveTheRow(
        string areaText,
        int row,
        bool isBelow,
        string? expectedBelow,
        string? expectedAbove
    )
    {
        Area area = Area.Parse(areaText);
        Assert.AreEqual(isBelow, area.SplitBelow(row, out Area? below, out Area? above));
        Assert.AreEqual(expectedBelow, below?.ToString());
        Assert.AreEqual(expectedAbove, above?.ToString());
    }

    [TestCase("A1:B1", 2, true, "A1", "B1")]
    [TestCase("A1", 2, true, "A1", null)]
    [TestCase("A1:C3", 1, false, null, "A1:C3")]
    public void SplitBeforeSplitsAreaIntoAreaToLeftAndRightOfColumn(
        string areaText,
        int column,
        bool isToLeft,
        string? expectedLeft,
        string? expectedRight
    )
    {
        Area area = Area.Parse(areaText);
        Assert.AreEqual(isToLeft, area.SplitBefore(column, out Area? left, out Area? right));
        Assert.AreEqual(expectedLeft, left?.ToString());
        Assert.AreEqual(expectedRight, right?.ToString());
    }

    [TestCase("A1:B1", 1, true, "B1", "A1")]
    [TestCase("B2:C3", 1, true, "B2:C3", null)]
    [TestCase("A1", 1, false, null, "A1")]
    public void SplitAfterSplitsAreaIntoAreaToRightAndLeftOfColumn(
        string areaText,
        int column,
        bool isToRight,
        string? expectedRight,
        string? expectedLeft
    )
    {
        Area area = Area.Parse(areaText);
        Assert.AreEqual(isToRight, area.SplitAfter(column, out Area? right, out Area? left));
        Assert.AreEqual(expectedRight, right?.ToString());
        Assert.AreEqual(expectedLeft, left?.ToString());
    }
}

using NUnit.Framework;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Cells;

[TestFixture]
public class SliceTests
{
    [Test]
    public void Stores_Values()
    {
        Slice<int> slice = new();
        Point point = new(574, 241);
        slice.Set(point, 1);
        Assert.AreEqual(1, slice[point]);
    }

    [Test]
    public void Setting_Value_To_Default_Clears_Element()
    {
        Slice<int> slice = new();
        Point point = new(574, 241);
        slice.Set(point, 1);
        Assert.AreEqual(574, slice.MaxRow);
        Assert.AreEqual(241, slice.MaxColumn);

        slice.Set(point, 0);

        Assert.AreEqual(0, slice.MaxRow);
        Assert.AreEqual(0, slice.MaxColumn);
    }

    [Test]
    public void Keeps_Track_Of_Max_Used_Coordinates()
    {
        Slice<int> slice = new();
        slice.Set(54, 32, 1);
        slice.Set(140, 32, 1);
        slice.Set(140, 72, 1);

        Assert.AreEqual(140, slice.MaxRow);
        Assert.AreEqual(72, slice.MaxColumn);

        slice.Set(140, 72, 0);

        Assert.AreEqual(140, slice.MaxRow);
        Assert.AreEqual(32, slice.MaxColumn);

        slice.Set(140, 32, 0);

        Assert.AreEqual(54, slice.MaxRow);
        Assert.AreEqual(32, slice.MaxColumn);

        slice.Set(54, 32, 0);

        Assert.AreEqual(0, slice.MaxRow);
        Assert.AreEqual(0, slice.MaxColumn);
    }

    [Test]
    public void Keeps_Track_Of_Used_Rows()
    {
        Slice<int> slice = new();
        Assert.IsEmpty(slice.UsedRows);

        slice.Set(new Point(1, 1), 1);
        CollectionAssert.AreEquivalent(new[] { 1 }, slice.UsedRows);

        slice.Set(new Point(70, 1), 1);
        CollectionAssert.AreEquivalent(new[] { 1, 70 }, slice.UsedRows);

        slice.Set(new Point(35, 1), 1);
        CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

        slice.Set(new Point(35, 2), 1);
        CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

        slice.Set(new Point(35, 1), 0);
        CollectionAssert.AreEquivalent(new[] { 1, 35, 70 }, slice.UsedRows);

        slice.Set(new Point(35, 2), 0);
        CollectionAssert.AreEquivalent(new[] { 1, 70 }, slice.UsedRows);

        slice.Set(new Point(1, 1), 0);
        CollectionAssert.AreEquivalent(new[] { 70 }, slice.UsedRows);

        slice.Set(new Point(70, 1), 0);
        Assert.IsEmpty(slice.UsedRows);
    }

    [Test]
    public void Keeps_Track_Of_Used_Columns()
    {
        Slice<int> slice = new();
        Assert.IsEmpty(slice.UsedColumns);

        slice.Set(new Point(1, 5), 1);
        CollectionAssert.AreEquivalent(new[] { 5 }, slice.UsedColumns);

        slice.Set(new Point(1, 750), 1);
        CollectionAssert.AreEquivalent(new[] { 5, 750 }, slice.UsedColumns);

        slice.Set(new Point(1, 90), 1);
        CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

        slice.Set(new Point(2, 5), 1);
        CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

        slice.Set(new Point(1, 5), 0);
        CollectionAssert.AreEquivalent(new[] { 5, 90, 750 }, slice.UsedColumns);

        slice.Set(new Point(2, 5), 0);
        CollectionAssert.AreEquivalent(new[] { 90, 750 }, slice.UsedColumns);

        slice.Set(new Point(1, 750), 0);
        CollectionAssert.AreEquivalent(new[] { 90 }, slice.UsedColumns);

        slice.Set(new Point(1, 90), 0);
        Assert.IsEmpty(slice.UsedColumns);
    }

    [Test]
    public void Clear_Range_Sets_Values_To_Default()
    {
        Slice<int> slice = new();
        Point outsideAddress = new(1, 1);
        slice.Set(outsideAddress, 1);
        Point firstCorner = new(50, 20);
        slice.Set(firstCorner, 1);
        Point insideAddress = new(55, 22);
        slice.Set(insideAddress, 1);
        Point lastCorner = new(60, 30);
        slice.Set(lastCorner, 1);

        slice.Clear(new Area(firstCorner, lastCorner));
        Assert.AreEqual(1, slice[outsideAddress]);
        Assert.AreEqual(0, slice[firstCorner]);
        Assert.AreEqual(0, slice[insideAddress]);
        Assert.AreEqual(0, slice[lastCorner]);
    }

    [Test]
    public void InsertAreaAndShiftDown_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
    {
        Slice<int> slice = new();
        slice.Set(1, 1, 1);
        slice.Set(3, 1, 2);
        Point purgedAddress = new(XLHelper.MaxRowNumber, 2);
        slice.Set(purgedAddress, 3);

        Point outsideAddress = new(1, 3);
        slice.Set(outsideAddress, 4);

        slice.InsertAreaAndShiftDown(new Area(new Point(1, 1), new Point(2, 2)));

        Assert.AreEqual(1, slice[3, 1]);
        Assert.AreEqual(2, slice[5, 1]);
        Assert.AreEqual(0, slice[XLHelper.MaxRowNumber, 2]);
        Assert.AreEqual(4, slice[outsideAddress]);
    }

    [Test]
    public void InsertAreaAndShiftRight_Moves_Area_Cells_Down_And_Purges_Values_Outside_Worksheet()
    {
        Slice<int> slice = new();
        slice.Set(1, 1, 1);
        slice.Set(1, 3, 2);
        Point purgedAddress = new(2, XLHelper.MaxColumnNumber);
        slice.Set(purgedAddress, 3);

        Point outsideAddress = new(3, 1);
        slice.Set(outsideAddress, 4);

        slice.InsertAreaAndShiftRight(new Area(new Point(1, 1), new Point(2, 2)));

        Assert.AreEqual(1, slice[1, 3]);
        Assert.AreEqual(2, slice[1, 5]);
        Assert.AreEqual(0, slice[purgedAddress]);
        Assert.AreEqual(4, slice[outsideAddress]);
    }

    [Test]
    public void DeleteAreaAndShiftUp_Moves_Area_Cells_Up()
    {
        Slice<int> slice = new();
        Point aboveAddress = new(1, 3);
        slice.Set(aboveAddress, 1);
        Point firstCorner = new(2, 2);
        slice.Set(firstCorner, 2);
        Point secondCorner = new(4, 5);
        slice.Set(secondCorner, 3);
        Point rightAddress = new(3, 6);
        slice.Set(rightAddress, 4);
        Point belowAddress = new(5, 3);
        slice.Set(belowAddress, 5);
        Point leftAddress = new(3, 1);
        slice.Set(leftAddress, 6);

        Area deleteArea = new(firstCorner, secondCorner);
        slice.DeleteAreaAndShiftUp(deleteArea);
        Assert.AreEqual(0, slice[firstCorner]);
        Assert.AreEqual(0, slice[secondCorner]);
        Assert.AreEqual(5, slice[belowAddress.Row - deleteArea.Height, belowAddress.Column]);
        Assert.AreEqual(1, slice[aboveAddress]);
        Assert.AreEqual(4, slice[rightAddress]);
        Assert.AreEqual(6, slice[leftAddress]);
    }

    [Test]
    public void DeleteAreaAndShiftLeft_Moves_Area_Cells_Left()
    {
        Slice<int> slice = new();
        Point leftAddress = new(3, 1);
        slice.Set(leftAddress, 1);
        Point firstCorner = new(2, 2);
        slice.Set(firstCorner, 2);
        Point secondCorner = new(5, 4);
        slice.Set(secondCorner, 3);
        Point belowAddress = new(6, 3);
        slice.Set(belowAddress, 4);
        Point rightAddress = new(3, 5);
        slice.Set(rightAddress, 5);
        Point aboveAddress = new(1, 3);
        slice.Set(aboveAddress, 6);

        Area deleteArea = new(firstCorner, secondCorner);
        slice.DeleteAreaAndShiftLeft(deleteArea);
        Assert.AreEqual(0, slice[firstCorner]);
        Assert.AreEqual(0, slice[secondCorner]);
        Assert.AreEqual(5, slice[rightAddress.Row, rightAddress.Column - deleteArea.Width]);
        Assert.AreEqual(1, slice[leftAddress]);
        Assert.AreEqual(4, slice[belowAddress]);
        Assert.AreEqual(6, slice[aboveAddress]);
    }

    [Test]
    public void DeleteAreaAndShiftUp_for_full_row_deletion_deletes_the_row_and_updates_column_usage()
    {
        // Test that a fast path for deletion of a area with full sheet width works the same
        // way as deletion of any area.
        Slice<int> slice = new();

        slice.Set(new Point(1, 3), 1);
        slice.Set(new Point(1, 7), 1);
        slice.Set(new Point(2, 4), 1);
        slice.Set(new Point(3, 6), 1);
        slice.Set(new Point(5, 4), 1);

        Assert.That(slice.MaxColumn, Is.EqualTo(7));
        Assert.That(slice.UsedColumns, Is.EquivalentTo([3, 4, 6, 7]));

        // Delete row with values at [3, 7]
        DeleteTopRow(6, [4, 6]);

        // Delete row with values at [4], but that is also used in the last row, so no change
        DeleteTopRow(6, [4, 6]);

        // Delete row with values at [6]
        DeleteTopRow(4, [4]);

        // Delete blank row, no change
        DeleteTopRow(4, [4]);

        // Delete last row [4], leaving the slice empty
        DeleteTopRow(0, []);
        return;

        void DeleteTopRow(int maxColumn, int[] columnUsage)
        {
            Area fullFirstRow = Area.Full.SliceFromTop(1);
            slice.DeleteAreaAndShiftUp(fullFirstRow);

            Assert.That(slice.MaxColumn, Is.EqualTo(maxColumn));
            Assert.That(slice.UsedColumns, Is.EquivalentTo(columnUsage));
        }
    }
}

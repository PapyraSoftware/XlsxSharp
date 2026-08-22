using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

/// <summary>
/// An immutable collection of areas. An equivalent of <c>ST_Sqref</c> (sequence of references).
/// List doesn't allow duplicate areas.
/// </summary>
internal class XLAreaList : IEnumerable<Area>
{
    internal static readonly XLAreaList Empty = new(new List<Area>());
    private readonly List<Area> _areas;

    internal XLAreaList(Area area)
    {
        this._areas = [area];
    }

    internal XLAreaList(List<Area> areas)
    {
        this._areas = areas;
    }

    internal int Count => this._areas.Count;

    internal Area this[int idx] => this._areas[idx];

    internal static XLAreaList FromRange(XLWorksheet worksheet, IXLRange range)
    {
        ThrowOnDifferentSheet(worksheet, range);
        return new XLAreaList(Area.FromRangeAddress(range.RangeAddress));
    }

    /// <exception cref="ArgumentException">Sequence is empty or a range is from a different sheet.</exception>
    internal static XLAreaList FromRanges(XLWorksheet worksheet, IEnumerable<IXLRange> value)
    {
        List<Area> areas = [];
        foreach (IXLRange range in value)
        {
            ThrowOnDifferentSheet(worksheet, range);
            areas.Add(Area.FromRangeAddress(range.RangeAddress));
        }

        if (areas.Count == 0)
        {
            throw new ArgumentException("Sequence is empty. At least one range is required.");
        }

        return new XLAreaList(areas);
    }

    internal XLAreaList With(Area area)
    {
        if (this._areas.Contains(area))
        {
            return this;
        }

        return new XLAreaList([.. this._areas, area]);
    }

    internal XLAreaList Without(Area area)
    {
        int indexToDelete = this._areas.IndexOf(area);
        if (indexToDelete == -1)
        {
            return this;
        }

        List<Area> newList = [.. this._areas];
        newList.RemoveAt(indexToDelete);
        return new XLAreaList(newList);
    }

    /// <summary>
    /// Insert and shift functionality as used in CF or DV.
    /// </summary>
    internal XLAreaList InsertAndShiftDown(Area insertedArea)
    {
        // Method is not symmetrical with InsertAndShiftRight, because the Excel doesn't produce
        // symmetrical results (e.g. original C3:E5 and insert down at C3 produces asymmetrical
        // results from insert right at E3).
        List<Area> result = new(this._areas.Count);
        foreach (Area originalArea in this._areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }

            // Skip all cases that don't shift or extend the area in some way.
            if (
                insertedArea.RightColumn < originalArea.LeftColumn
                || insertedArea.LeftColumn > originalArea.RightColumn
                || insertedArea.TopRow > originalArea.BottomRow + 1
            )
            {
                result.Add(originalArea);
                continue;
            }

            if (
                originalArea.SplitAbove(insertedArea.TopRow, out Area? above, out Area? remaining)
                && above.Value.LeftColumn >= insertedArea.LeftColumn
                && above.Value.RightColumn <= insertedArea.RightColumn
            )
            {
                // Special case: If inserted area is to the full width of original area and there is something above,
                // the whole area is just extended downwards. The optional null check is there if inserted area
                // attaches to the bottom of the original area.
                Area mergedAndExtended = above.Value.ExtendBelow(
                    insertedArea.Height + (remaining?.Height ?? 0)
                );
                result.Add(mergedAndExtended);
                continue;
            }

            Area? left = null,
                right = null;
            if (remaining is not null)
            {
                remaining.Value.SplitBefore(insertedArea.LeftColumn, out left, out remaining);
            }

            if (remaining is not null)
            {
                remaining.Value.SplitAfter(insertedArea.RightColumn, out right, out remaining);
            }

            if (above is not null)
            {
                result.Add(above.Value);
            }

            if (left is not null)
            {
                result.Add(left.Value);
            }

            if (right is not null)
            {
                result.Add(right.Value);
            }

            if (above is not null)
            {
                // There was something above the inserted area so extend
                if (remaining is not null)
                {
                    Area extended = remaining.Value.ExtendBelow(insertedArea.Height);
                    result.Add(extended);
                }
                else if (insertedArea.TopRow == originalArea.BottomRow + 1)
                {
                    // Attaches partial cover at the bottom of original area, e.g. insert to B2 with original A1:C1
                    Area cutToWidth = new(
                        insertedArea.TopRow,
                        Math.Max(insertedArea.LeftColumn, originalArea.LeftColumn),
                        insertedArea.BottomRow,
                        Math.Min(insertedArea.RightColumn, originalArea.RightColumn)
                    );
                    result.Add(cutToWidth);
                }
            }
            else
            {
                // There was nothing above the inserted area, so shift.
                if (remaining is null)
                {
                    throw new UnreachableException();
                }

                if (remaining.Value.ShiftRowsAndClip(insertedArea.Height) is { } shifted)
                {
                    result.Add(shifted);
                }
            }
        }

        return new XLAreaList(result);
    }

    /// <summary>
    /// Insert and shift functionality as used in CF or DV.
    /// </summary>
    internal XLAreaList InsertAndShiftRight(Area insertedArea)
    {
        // Method is not symmetrical with InsertAndShiftDown, because the Excel doesn't produce
        // symmetrical results (e.g. original C3:E5 and insert down at C3 produces asymmetrical
        // results from insert right at E3).
        List<Area> result = new(this._areas.Count);
        foreach (Area originalArea in this._areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            // Skip all cases that don't shift or extend the area in some way.
            if (
                insertedArea.BottomRow < originalArea.TopRow
                || insertedArea.TopRow > originalArea.BottomRow
                || insertedArea.LeftColumn > originalArea.RightColumn + 1
            )
            {
                result.Add(originalArea);
                continue;
            }

            // Deal with special case of attachment at the right side
            if (insertedArea.LeftColumn == originalArea.RightColumn + 1)
            {
                if (
                    originalArea.TopRow >= insertedArea.TopRow
                    && originalArea.BottomRow <= insertedArea.BottomRow
                )
                {
                    result.Add(originalArea.ExtendRight(insertedArea.Width));
                }
                else
                {
                    // Attaches at the right of original area, e.g. insert to B2 with original A1:C1
                    Area cutToHeight = new(
                        Math.Max(insertedArea.TopRow, originalArea.TopRow),
                        insertedArea.LeftColumn,
                        Math.Min(insertedArea.BottomRow, originalArea.BottomRow),
                        insertedArea.RightColumn
                    );
                    result.Add(originalArea);
                    result.Add(cutToHeight);
                }

                continue;
            }

            Area? below = null,
                left = null;
            originalArea.SplitAbove(insertedArea.TopRow, out Area? above, out Area? remaining);

            if (remaining is not null)
            {
                remaining.Value.SplitBelow(insertedArea.BottomRow, out below, out remaining);
            }

            if (remaining is not null)
            {
                remaining.Value.SplitBefore(insertedArea.LeftColumn, out left, out remaining);
            }

            // There must be something. We know that inserted area intersects original area (we took care of special
            // case of right side attachment in an if above) and we only cut three times on each side of
            // the intersection, so something must be left.
            if (remaining is null)
            {
                throw new UnreachableException();
            }

            if (above is not null)
            {
                result.Add(above.Value);
            }

            if (below is not null)
            {
                result.Add(below.Value);
            }

            if (left is not null)
            {
                // There was something on the left of the inserted area so extend
                Area mergedAndExtended = left.Value.ExtendRight(
                    insertedArea.Width + remaining.Value.Width
                );
                result.Add(mergedAndExtended);
            }
            else
            {
                // There is nothing on the left side, so shift
                if (remaining.Value.ShiftColumnsAndClip(insertedArea.Width) is { } shifted)
                {
                    result.Add(shifted);
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftUp(Area deletedArea)
    {
        Area groove = deletedArea.ExtendBelow(XlsxSharp.XLHelper.MaxRowNumber);
        List<Area> result = new(this._areas.Count);
        foreach (Area originalArea in this._areas)
        {
            if (originalArea.HasFullColumnHeight)
            {
                result.Add(originalArea);
                continue;
            }
            bool deleteWontSplitOriginalArea =
                deletedArea.LeftColumn <= originalArea.LeftColumn
                && deletedArea.RightColumn >= originalArea.RightColumn;
            if (deleteWontSplitOriginalArea)
            {
                Area? shiftedArea = originalArea.ShiftOrShrinkUp(
                    deletedArea.TopRow,
                    deletedArea.Height
                );
                if (shiftedArea is not null)
                {
                    result.Add(shiftedArea.Value);
                }
            }
            else
            {
                Area? inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it upwards
                    Area? shiftedArea = inGrooveArea.Value.ShiftOrShrinkUp(
                        deletedArea.TopRow,
                        deletedArea.Height
                    );
                    if (shiftedArea is not null)
                    {
                        result.Add(shiftedArea.Value);
                    }
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteAndShiftLeft(Area deletedArea)
    {
        Area groove = deletedArea.ExtendRight(XlsxSharp.XLHelper.MaxColumnNumber);
        List<Area> result = new(this._areas.Count);
        foreach (Area originalArea in this._areas)
        {
            if (originalArea.HasFullRowWidth)
            {
                result.Add(originalArea);
                continue;
            }

            bool deleteWontSplitOriginalArea =
                deletedArea.TopRow <= originalArea.TopRow
                && deletedArea.BottomRow >= originalArea.BottomRow;
            if (deleteWontSplitOriginalArea)
            {
                Area? shiftedArea = originalArea.ShiftOrShrinkLeft(
                    deletedArea.LeftColumn,
                    deletedArea.Width
                );
                if (shiftedArea is not null)
                {
                    result.Add(shiftedArea.Value);
                }
            }
            else
            {
                Area? inGrooveArea = originalArea.Exclude(groove, result);
                if (inGrooveArea is not null)
                {
                    // There is something to shift, so shift it leftward
                    Area? shiftedArea = inGrooveArea.Value.ShiftOrShrinkLeft(
                        deletedArea.LeftColumn,
                        deletedArea.Width
                    );
                    if (shiftedArea is not null)
                    {
                        result.Add(shiftedArea.Value);
                    }
                }
            }
        }

        return new XLAreaList(result);
    }

    internal XLAreaList DeleteWithoutShift(Area deletedArea)
    {
        List<Area> result = new(this._areas.Count);
        foreach (Area originalArea in this._areas)
        {
            originalArea.Exclude(deletedArea, result);
        }

        return new XLAreaList(result);
    }

    internal XLAreaList GetConsolidated()
    {
        return XLRangeConsolidationEngine.Consolidate(this);
    }

    internal bool IntersectsWith(Area otherArea)
    {
        foreach (Area area in this._areas)
        {
            if (area.Intersects(otherArea))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Return areas in the list (with the original size) intersecting with the <paramref name="otherArea"/>.
    /// </summary>
    internal IEnumerable<Area> IntersectingWith(Area otherArea)
    {
        foreach (Area area in this._areas)
        {
            if (area.Intersects(otherArea))
            {
                yield return area;
            }
        }
    }

    /// <summary>
    /// A helper function used mostly in copy&amp;paste functionality. It takes the areas,
    /// intersects them with the <paramref name="areaToCopy"/> and shifts it to the <paramref name="target"/>.
    /// If there are areas, return it in the <paramref name="result"/>.
    /// </summary>
    internal bool TryCopyAreaTo(
        Point target,
        Area areaToCopy,
        [NotNullWhen(true)] out XLAreaList? result
    )
    {
        int rowShift = target.Row - areaToCopy.FirstPoint.Row;
        int columnShift = target.Column - areaToCopy.FirstPoint.Column;
        List<Area>? copyList = null;
        foreach (Area area in this._areas)
        {
            if (area.Intersect(areaToCopy) is not { } intersection)
            {
                continue;
            }

            // End can but cut off, but the area will always have at least 1x1 so it is valid
            if (intersection.ShiftAndClip(rowShift, columnShift) is not { } shiftedArea)
            {
                continue;
            }

            copyList ??= [];
            copyList.Add(shiftedArea);
        }

        if (copyList is not null)
        {
            result = new XLAreaList(copyList);
            return true;
        }

        result = null;
        return false;
    }

    internal XLAreaList Excluding(Area excludedArea)
    {
        if (!this.IntersectsWith(excludedArea))
        {
            return this;
        }

        List<Area> list = [];
        foreach (Area area in this._areas)
        {
            area.Exclude(excludedArea, list);
        }

        return new XLAreaList(list);
    }

    public IEnumerator<Area> GetEnumerator()
    {
        return this._areas.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    internal string ToSpaceList()
    {
        return string.Join(" ", this._areas);
    }

    private static void ThrowOnDifferentSheet(XLWorksheet worksheet, IXLRange range)
    {
        if (range.Worksheet is not null && range.Worksheet != worksheet)
        {
            throw new ArgumentException(
                $"Range {range} belongs to worksheet {range.Worksheet.Name}, but must be from worksheet {worksheet.Name}."
            );
        }
    }
}

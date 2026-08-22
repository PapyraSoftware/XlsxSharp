using System;
using System.Collections;
using System.Collections.Generic;

namespace XlsxSharp.Excel;

/// <summary>
/// A specification of an area (rectangular range) of a sheet.
/// </summary>
internal readonly struct SheetArea : IEquatable<SheetArea>, IEnumerable<SheetPoint>
{
    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public readonly string Name;

    /// <summary>
    /// An area in the sheet.
    /// </summary>
    public readonly Area Area;

    public SheetArea(String name, Area area)
    {
        if (string.IsNullOrEmpty(name))
        {
            throw new ArgumentException(nameof(name));
        }

        this.Name = name;
        this.Area = area;
    }

    public static bool operator ==(SheetArea lhs, SheetArea rhs) => lhs.Equals(rhs);

    public static bool operator !=(SheetArea lhs, SheetArea rhs) => !(lhs == rhs);

    public IEnumerator<SheetPoint> GetEnumerator()
    {
        for (int row = this.Area.TopRow; row <= this.Area.BottomRow; ++row)
        {
            for (int col = this.Area.LeftColumn; col <= this.Area.RightColumn; ++col)
            {
                yield return new SheetPoint(this.Name, row, col);
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    internal static SheetArea From(IXLRange range)
    {
        if (range.Worksheet is null)
        {
            throw new ArgumentException("Range doesn't contain sheet.", nameof(range));
        }

        return new SheetArea(range.Worksheet.Name, Area.FromRangeAddress(range.RangeAddress));
    }

    internal static SheetArea From(XLRangeAddress address)
    {
        if (address.Worksheet is null)
        {
            throw new ArgumentException("Range doesn't contain sheet.", nameof(address));
        }

        return new SheetArea(address.Worksheet.Name, Area.FromRangeAddress(address));
    }

    public bool Equals(SheetArea other)
    {
        return this.Area == other.Area
            && XlsxSharp.XLHelper.SheetComparer.Equals(this.Name, other.Name);
    }

    public override bool Equals(object? obj)
    {
        return obj is SheetArea other && this.Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (XlsxSharp.XLHelper.SheetComparer.GetHashCode(this.Name) * 397)
                ^ this.Area.GetHashCode();
        }
    }

    /// <summary>
    /// Perform an intersection.
    /// </summary>
    /// <param name="other">The area that is being intersected with this one.</param>
    /// <returns>The intersection (=same sheet and has non-empty intersection) or null if intersection isn't possible.</returns>
    public SheetArea? Intersect(SheetArea other)
    {
        if (!XlsxSharp.XLHelper.SheetComparer.Equals(this.Name, other.Name))
        {
            return null;
        }

        Area? intersectionRange = this.Area.Intersect(other.Area);
        if (intersectionRange is null)
        {
            return null;
        }

        return new SheetArea(this.Name, intersectionRange.Value);
    }

    public void Deconstruct(out string sheetName, out Area area)
    {
        sheetName = this.Name;
        area = this.Area;
    }

    public override string ToString()
    {
        return $"{this.Name}!{this.Area}";
    }
}

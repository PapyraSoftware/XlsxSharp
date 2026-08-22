#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.Patterns;

/// <summary>
/// Implementation of QuadTree adapted to Excel worksheet specifics. Differences with the classic implementation
/// are that the topmost level is split to 128 square parts (2 columns of 64 blocks, each 8192*8192 cells) and that splitting
/// the quadrant onto 4 smaller quadrants does not depend on the number of items in this quadrant. When the range is added to the
/// QuadTree it is placed on the bottommost level where it fits to a single quadrant. That means, row-wide and column-wide ranges
/// are always placed at the level 0, and the smaller the range is the deeper it goes down the tree. This approach eliminates
/// the need of transferring ranges between levels.
/// </summary>
internal class Quadrant
{
    #region Public Properties

    /// <summary>
    /// Smaller quadrants which the current one is split to. Is NULL until ranges are added to child quadrants.
    /// </summary>
    public IReadOnlyList<Quadrant> Children { get; private set; }

    /// <summary>
    /// The level of current quadrant. Top most has level 0, child quadrants has levels (Level + 1).
    /// </summary>
    public byte Level { get; }

    /// <summary>
    /// Minimum column included in this quadrant.
    /// </summary>
    public int MinimumColumn { get; }

    /// <summary>
    /// Minimum row included in this quadrant.
    /// </summary>
    public int MinimumRow { get; }

    /// <summary>
    /// Maximum column included in this quadrant.
    /// </summary>
    public int MaximumColumn { get; }

    /// <summary>
    /// Maximum row included in this quadrant.
    /// </summary>
    public int MaximumRow { get; }

    /// <summary>
    /// Collection of ranges belonging to this quadrant (does not include ranges from child quadrants).
    /// </summary>
    public IEnumerable<IXLAddressable> Ranges
    {
        get => this._ranges?.Values.AsEnumerable();
    }

    /// <summary>
    /// The number of current quadrant by horizontal axis.
    /// </summary>
    public short X { get; private set; }

    /// <summary>
    /// The number of current quadrant by vertical axis.
    /// </summary>
    public short Y { get; private set; }

    #endregion Public Properties

    #region Constructors

    public Quadrant()
        : this(0, 0, 0) { }

    private Quadrant(byte level, short x, short y)
    {
        this.Level = level;
        this.X = x;
        this.Y = y;

        this.MinimumColumn =
            (this.Level == 0)
                ? 1
                : 1 + XlsxSharp.XLHelper.MaxColumnNumber / (int)Math.Pow(2, this.Level) * this.X;
        this.MinimumRow =
            (this.Level == 0)
                ? 1
                : 1 + XlsxSharp.XLHelper.MaxColumnNumber / (int)Math.Pow(2, this.Level) * this.Y; //MaxColumnNumber here is not a mistake
        this.MaximumColumn =
            (this.Level == 0)
                ? XlsxSharp.XLHelper.MaxColumnNumber
                : XlsxSharp.XLHelper.MaxColumnNumber / (int)Math.Pow(2, this.Level) * (this.X + 1);
        this.MaximumRow =
            (this.Level == 0)
                ? XlsxSharp.XLHelper.MaxRowNumber
                : XlsxSharp.XLHelper.MaxColumnNumber / (int)Math.Pow(2, this.Level) * (this.Y + 1); //MaxColumnNumber here is not a mistake
    }

    #endregion Constructors

    #region Public Methods

    /// <summary>
    /// Add a range to the quadrant or to one of the child quadrants (recursively).
    /// </summary>
    /// <returns>True, if range was successfully added, false if it has been added before.</returns>
    public bool Add(IXLAddressable range)
    {
        bool res = false;
        IReadOnlyList<Quadrant> children = this.Children ?? this.CreateChildren().ToList();
        bool addToChild = false;
        foreach (Quadrant childQuadrant in children)
        {
            IXLRangeAddress rangeAddress = range.RangeAddress;
            if (childQuadrant.Covers(in rangeAddress))
            {
                res |= childQuadrant.Add(range);
                addToChild = true;
                break;
            }
        }

        if (!addToChild)
        {
            res = this.AddInternal(range);
        }

        if (this.Children == null && addToChild)
        {
            this.Children = children;
        }

        return res;
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively).
    /// </summary>
    public IEnumerable<IXLAddressable> GetAll()
    {
        if (this.Ranges != null)
        {
            foreach (IXLAddressable range in this.Ranges)
            {
                yield return range;
            }
        }

        if (this.Children != null)
        {
            foreach (Quadrant childQuadrant in this.Children)
            {
                IEnumerable<IXLAddressable> childRanges = childQuadrant.GetAll();
                foreach (IXLAddressable range in childRanges)
                {
                    yield return range;
                }
            }
        }
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively) that intersect the specified address.
    /// </summary>
    public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLRangeAddress rangeAddress)
    {
        if (this.Ranges != null)
        {
            foreach (IXLAddressable range in this.Ranges)
            {
                if (range.RangeAddress.Intersects(rangeAddress))
                {
                    yield return range;
                }
            }
        }

        if (this.Children != null)
        {
            foreach (Quadrant childQuadrant in this.Children)
            {
                if (childQuadrant.Intersects(in rangeAddress))
                {
                    IEnumerable<IXLAddressable> childRanges = childQuadrant.GetIntersectedRanges(
                        rangeAddress
                    );
                    foreach (IXLAddressable range in childRanges)
                    {
                        yield return range;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively) that cover the specified address.
    /// </summary>
    public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLAddress address)
    {
        if (this.Ranges != null)
        {
            foreach (IXLAddressable range in this.Ranges)
            {
                if (range.RangeAddress.Contains(address))
                {
                    yield return range;
                }
            }
        }

        if (this.Children != null)
        {
            foreach (Quadrant childQuadrant in this.Children)
            {
                if (childQuadrant.Covers(in address))
                {
                    IEnumerable<IXLAddressable> childRanges = childQuadrant.GetIntersectedRanges(
                        address
                    );
                    foreach (IXLAddressable range in childRanges)
                    {
                        yield return range;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Remove the range from the quadrant or from child quadrants (recursively).
    /// </summary>
    /// <returns>True if the range was removed, false if it does not exist in the QuadTree.</returns>
    public bool Remove(IXLRangeAddress rangeAddress)
    {
        bool res = false;

        bool coveredByChild = false;
        if (this.Children != null)
        {
            foreach (Quadrant childQuadrant in this.Children)
            {
                if (childQuadrant.Covers(rangeAddress))
                {
                    res |= childQuadrant.Remove(rangeAddress);
                    coveredByChild = true;
                }
            }
        }

        if (!coveredByChild)
        {
            if (this._ranges?.Remove(rangeAddress) == true)
            {
                res = true;
            }
        }

        return res;
    }

    /// <summary>
    /// Remove all the ranges matching specified criteria from the quadrant and its child quadrants (recursively).
    /// Don't use it for searching intersections as it would be much less efficient than <see cref="GetIntersectedRanges(IXLRangeAddress)"/>.
    /// </summary>
    public IEnumerable<IXLAddressable> RemoveAll(Predicate<IXLAddressable> predicate)
    {
        if (this._ranges != null)
        {
            IEnumerable<IXLAddressable> ranges = this._ranges.Values.Where(r => predicate(r));
            List<IXLRangeAddress> keysToRemove = [];
            foreach (IXLAddressable range in ranges)
            {
                keysToRemove.Add(range.RangeAddress);
                yield return range;
            }

            foreach (IXLRangeAddress keyToRemove in keysToRemove)
            {
                this._ranges.Remove(keyToRemove);
            }
        }

        if (this.Children != null)
        {
            foreach (Quadrant childQuadrant in this.Children)
            foreach (IXLAddressable childRange in childQuadrant.RemoveAll(predicate))
            {
                yield return childRange;
            }
        }
    }

    #endregion Public Methods

    #region Private Fields

    /// <summary>
    /// Maximum depth of the QuadTree. Value 10 corresponds to the smallest quadrants having size 16*16 cells.
    /// </summary>
    private const byte MAX_LEVEL = 10;

    /// <summary>
    /// Collection of ranges belonging to the current quadrant (that cannot fit into child quadrants).
    /// </summary>
    private Dictionary<IXLRangeAddress, IXLAddressable> _ranges;

    #endregion Private Fields

    #region Private Methods

    /// <summary>
    /// Add a range to the collection of quadrant's own ranges.
    /// </summary>
    /// <returns>True if the range was successfully added, false if it had been added before.</returns>
    private bool AddInternal(IXLAddressable range)
    {
        if (this._ranges == null)
        {
            this._ranges = new Dictionary<IXLRangeAddress, IXLAddressable>();
        }

        if (this._ranges.ContainsKey(range.RangeAddress))
        {
            return false;
        }

        this._ranges.Add(range.RangeAddress, range);
        return true;
    }

    /// <summary>
    /// Check if the current quadrant fully covers the specified address.
    /// </summary>
    private bool Covers(in IXLRangeAddress rangeAddress)
    {
        return this.MinimumColumn <= rangeAddress.FirstAddress.ColumnNumber
            && this.MaximumColumn >= rangeAddress.LastAddress.ColumnNumber
            && this.MinimumRow <= rangeAddress.FirstAddress.RowNumber
            && this.MaximumRow >= rangeAddress.LastAddress.RowNumber;
    }

    /// <summary>
    /// Check if the current quadrant covers the specified address.
    /// </summary>
    private bool Covers(in IXLAddress address)
    {
        return this.MinimumColumn <= address.ColumnNumber
            && this.MaximumColumn >= address.ColumnNumber
            && this.MinimumRow <= address.RowNumber
            && this.MaximumRow >= address.RowNumber;
    }

    /// <summary>
    /// Check if the current quadrant intersects the specified address.
    /// </summary>
    private bool Intersects(in IXLRangeAddress rangeAddress)
    {
        return (
                (
                    this.MinimumRow <= rangeAddress.FirstAddress.RowNumber
                    && rangeAddress.FirstAddress.RowNumber <= this.MaximumRow
                )
                || (
                    rangeAddress.FirstAddress.RowNumber <= this.MinimumRow
                    && this.MinimumRow <= rangeAddress.LastAddress.RowNumber
                )
            )
            && (
                (
                    this.MinimumColumn <= rangeAddress.FirstAddress.ColumnNumber
                    && rangeAddress.FirstAddress.ColumnNumber <= this.MaximumColumn
                )
                || (
                    rangeAddress.FirstAddress.ColumnNumber <= this.MinimumColumn
                    && this.MinimumColumn <= rangeAddress.LastAddress.ColumnNumber
                )
            );
    }

    /// <summary>
    /// Create a collection of child quadrants dividing the current one.
    /// </summary>
    private IEnumerable<Quadrant> CreateChildren()
    {
        byte childLevel = (byte)(this.Level + 1);
        if (childLevel > MAX_LEVEL)
        {
            yield break;
        }

        byte xCount = 2; // Always divide on halves
        byte yCount = (byte)(
            (this.Level == 0)
                ? (XlsxSharp.XLHelper.MaxRowNumber / XlsxSharp.XLHelper.MaxColumnNumber)
                : 2
        ); // Level 0 divide onto 64 parts, the rest - on halves

        for (byte dy = 0; dy < yCount; dy++)
        {
            for (byte dx = 0; dx < xCount; dx++)
            {
                yield return new Quadrant(
                    childLevel,
                    (short)(this.X * 2 + dx),
                    (short)(this.Y * 2 + dy)
                );
            }
        }
    }

    #endregion Private Methods
}

/// <summary>
/// A generic version of <see cref="Quadrant"/>
/// </summary>
internal class Quadrant<T> : Quadrant
    where T : IXLAddressable
{
    public new IEnumerable<T> Ranges => base.Ranges.Cast<T>();

    public bool Add(T range)
    {
        return base.Add(range);
    }

    public new IEnumerable<T> GetAll()
    {
        return base.GetAll().Cast<T>();
    }

    public new IEnumerable<T> GetIntersectedRanges(IXLRangeAddress rangeAddress)
    {
        return base.GetIntersectedRanges(rangeAddress).Cast<T>();
    }

    public new IEnumerable<T> GetIntersectedRanges(IXLAddress address)
    {
        return base.GetIntersectedRanges(address).Cast<T>();
    }

    public bool Remove(T range)
    {
        return this.Remove(range.RangeAddress);
    }

    public IEnumerable<T> RemoveAll(Predicate<T> predicate)
    {
        return base.RemoveAll(r => predicate((T)r)).Cast<T>();
    }
}

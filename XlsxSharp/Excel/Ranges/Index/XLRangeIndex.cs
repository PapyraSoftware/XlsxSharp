using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Patterns;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel.Index;

/// <summary>
/// Implementation of <see cref="IXLRangeIndex"/> internally using QuadTree.
/// </summary>
internal abstract class XLRangeIndex : IXLRangeIndex
{
    #region Public Constructors

    public XLRangeIndex(IXLWorksheet worksheet)
    {
        this._worksheet = worksheet;
        this._rangeList = [];
        ((XLWorksheet)this._worksheet).RegisterRangeIndex(this);
    }

    #endregion Public Constructors

    #region Public Methods

    public abstract bool MatchesType(XLRangeType rangeType);

    public bool Add(IXLAddressable range)
    {
        ArgumentNullException.ThrowIfNull(range);

        if (!range.RangeAddress.IsValid)
        {
            throw new ArgumentException("Range is invalid");
        }

        this.CheckWorksheet(range.RangeAddress.Worksheet);

        this._count++;
        if (this._count < MinimumCountForIndexing)
        {
            if (this._rangeList.Any(r => r == range))
            {
                return false;
            }

            this._rangeList.Add(range);
            return true;
        }

        if (this._quadTree == null)
        {
            this.InitializeTree();
        }

        return this._quadTree!.Add(range);
    }

    public bool Contains(in XLAddress address)
    {
        this.CheckWorksheet(address.Worksheet);

        if (this._quadTree == null)
        {
            XLAddress addr = address;
            return this._rangeList.Any(r => r.RangeAddress.Contains(addr));
        }

        return this._quadTree.GetIntersectedRanges(address).Any();
    }

    public IEnumerable<IXLAddressable> GetAll()
    {
        if (this._quadTree == null)
        {
            return this._rangeList.AsEnumerable();
        }

        return this._quadTree.GetAll();
    }

    public IEnumerable<IXLAddressable> GetIntersectedRanges(XLRangeAddress rangeAddress)
    {
        this.CheckWorksheet(rangeAddress.Worksheet);

        if (this._quadTree == null)
        {
            return this._rangeList.Where(r => r.RangeAddress.Intersects(rangeAddress));
        }

        return this._quadTree.GetIntersectedRanges(rangeAddress);
    }

    public IEnumerable<IXLAddressable> GetIntersectedRanges(XLAddress address)
    {
        this.CheckWorksheet(address.Worksheet);

        if (this._quadTree == null)
        {
            return this._rangeList.Where(r => r.RangeAddress.Contains(address));
        }

        return this._quadTree.GetIntersectedRanges(address);
    }

    public bool Intersects(in XLRangeAddress rangeAddress)
    {
        this.CheckWorksheet(rangeAddress.Worksheet);

        if (this._quadTree == null)
        {
            XLRangeAddress addr = rangeAddress;
            return this._rangeList.Any(r => r.RangeAddress.Intersects(addr));
        }

        return this._quadTree.GetIntersectedRanges(rangeAddress).Any();
    }

    public bool Remove(IXLRangeAddress rangeAddress)
    {
        ArgumentNullException.ThrowIfNull(rangeAddress);

        this.CheckWorksheet(rangeAddress.Worksheet);

        if (this._quadTree == null)
        {
            return this._rangeList.RemoveAll(r => Equals(r.RangeAddress, rangeAddress)) > 0;
        }

        return this._quadTree.Remove(rangeAddress);
    }

    public int RemoveAll(Predicate<IXLAddressable>? predicate = null)
    {
        predicate = predicate ?? (_ => true);

        if (this._quadTree == null)
        {
            return this._rangeList.RemoveAll(predicate);
        }

        return this._quadTree.RemoveAll(predicate).Count();
    }

    #endregion Public Methods

    #region Private Fields

    /// <summary>
    /// The minimum number of ranges to be included into a QuadTree. Until it is reached the ranges
    /// are added into a simple list to minimize the overhead of searching intersections on small collections.
    /// </summary>
    private const int MinimumCountForIndexing = 20;

    /// <summary>
    /// A collection of ranges used before the QuadTree is initialized (until <see cref="MinimumCountForIndexing"/>
    /// is reached.
    /// </summary>
    protected readonly List<IXLAddressable> _rangeList;

    private readonly IXLWorksheet _worksheet;
    private int _count = 0;
    protected Quadrant? _quadTree;

    #endregion Private Fields

    #region Private Methods

    private void CheckWorksheet(IXLWorksheet? worksheet)
    {
        if (worksheet != this._worksheet)
        {
            throw new ArgumentException("Range belongs to a different worksheet");
        }
    }

    private void InitializeTree()
    {
        this._quadTree = this.CreateQuadTree();
        this._rangeList.ForEach(r => this._quadTree.Add(r));
        this._rangeList.Clear();
    }

    protected virtual Quadrant CreateQuadTree() => new();

    #endregion Private Methods
}

/// <summary>
/// Generic version of <see cref="XLRangeIndex"/>.
/// </summary>
internal class XLRangeIndex<T> : XLRangeIndex, IXLRangeIndex<T>
    where T : IXLAddressable
{
    public XLRangeIndex(IXLWorksheet worksheet)
        : base(worksheet) { }

    public bool Add(T range) => base.Add(range);

    public int RemoveAll(Predicate<T>? predicate)
    {
        predicate = predicate ?? (_ => true);

        return base.RemoveAll(r => predicate((T)r));
    }

    public new IEnumerable<T> GetIntersectedRanges(XLRangeAddress rangeAddress) =>
        base.GetIntersectedRanges(rangeAddress).Cast<T>();

    public new IEnumerable<T> GetIntersectedRanges(XLAddress address) =>
        base.GetIntersectedRanges(address).Cast<T>();

    public override bool MatchesType(XLRangeType rangeType)
    {
        Type innerType = typeof(T);

        if (innerType == typeof(IXLRangeBase) || innerType == typeof(XLRangeBase))
        {
            return true;
        }

        switch (rangeType)
        {
            case XLRangeType.Range:
                return innerType == typeof(IXLRange) || innerType == typeof(XLRange);

            case XLRangeType.Column:
                return innerType == typeof(IXLColumn) || innerType == typeof(XLColumn);

            case XLRangeType.Row:
                return innerType == typeof(IXLRow) || innerType == typeof(XLRow);

            case XLRangeType.RangeColumn:
                return innerType == typeof(IXLRangeColumn) || innerType == typeof(XLRangeColumn);

            case XLRangeType.RangeRow:
                return innerType == typeof(IXLRangeRow) || innerType == typeof(XLRangeRow);

            case XLRangeType.Table:
                return innerType == typeof(IXLTable) || innerType == typeof(XLTable);

            case XLRangeType.Worksheet:
                return innerType == typeof(IXLWorksheet) || innerType == typeof(XLWorksheet);

            default:
                throw new NotImplementedException(nameof(rangeType));
        }
    }

    public new IEnumerable<T> GetAll() => base.GetAll().Cast<T>();

    protected override Quadrant CreateQuadTree() => new Quadrant<T>();
}

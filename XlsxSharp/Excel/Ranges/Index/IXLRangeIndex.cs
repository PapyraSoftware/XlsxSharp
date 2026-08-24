namespace XlsxSharp.Excel.Index;

/// <summary>
/// Interface for the engine aimed to speed-up the search for the range intersections.
/// </summary>
internal interface IXLRangeIndex
{
    public bool Add(IXLAddressable range);

    public bool Remove(IXLRangeAddress rangeAddress);

    public int RemoveAll(Predicate<IXLAddressable>? predicate = null);

    public IEnumerable<IXLAddressable> GetIntersectedRanges(XLRangeAddress rangeAddress);

    public IEnumerable<IXLAddressable> GetIntersectedRanges(XLAddress address);

    public IEnumerable<IXLAddressable> GetAll();

    public bool Intersects(in XLRangeAddress rangeAddress);

    public bool Contains(in XLAddress address);

    public bool MatchesType(XLRangeType rangeType);
}

internal interface IXLRangeIndex<T> : IXLRangeIndex
    where T : IXLAddressable
{
    public bool Add(T range);

    public int RemoveAll(Predicate<T>? predicate = null);

    public new IEnumerable<T> GetIntersectedRanges(XLRangeAddress rangeAddress);

    public new IEnumerable<T> GetIntersectedRanges(XLAddress address);

    public new IEnumerable<T> GetAll();
}

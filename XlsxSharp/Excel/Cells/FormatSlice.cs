using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

internal class FormatSlice : ISlice
{
    private readonly Slice<XLCellFormatValue?> _slice = new();

    public bool IsEmpty => this._slice.IsEmpty;

    public int MaxColumn => this._slice.MaxColumn;

    public int MaxRow => this._slice.MaxRow;

    public Dictionary<int, int>.KeyCollection UsedColumns => this._slice.UsedColumns;

    public IEnumerable<int> UsedRows => this._slice.UsedRows;

    public void Clear(Area area) => this._slice.Clear(area);

    public void DeleteAreaAndShiftLeft(Area areaToDelete) =>
        this._slice.DeleteAreaAndShiftLeft(areaToDelete);

    public void DeleteAreaAndShiftUp(Area areaToDelete) =>
        this._slice.DeleteAreaAndShiftUp(areaToDelete);

    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false) =>
        this._slice.GetEnumerator(area, reverse);

    public void InsertAreaAndShiftDown(Area areaToInsert) =>
        this._slice.InsertAreaAndShiftDown(areaToInsert);

    public void InsertAreaAndShiftRight(Area areaToInsert) =>
        this._slice.InsertAreaAndShiftRight(areaToInsert);

    public bool IsUsed(Point address) => this._slice.IsUsed(address);

    public void Swap(Point sp1, Point sp2) => this._slice.Swap(sp1, sp2);

    public void Set(Point point, XLCellFormatValue? value) => this._slice.Set(point, value);

    internal void SetAll(Area area, XLCellFormatValue? value) => this._slice.SetAll(area, value);

    internal XLCellFormatValue? GetFormat(Point point) => this._slice[point];

    internal void AddUsedFormat(HashSet<XLCellFormatValue> usedCellFormats)
    {
        IEnumerator<Point> enumerator = this.GetEnumerator(Area.Full);
        while (enumerator.MoveNext())
        {
            if (this._slice[enumerator.Current] is { } format)
            {
                usedCellFormats.Add(format);
            }
        }
    }
}

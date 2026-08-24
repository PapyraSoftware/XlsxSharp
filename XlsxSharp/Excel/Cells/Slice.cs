#nullable disable

using System.Collections;
using System.Diagnostics;

namespace XlsxSharp.Excel;

/// <summary>
/// Slice is a sparse array that stores a part of cell information (e.g. only values,
/// only styles ...). Slice has same size as a worksheet. If some cells are pushed out
/// of the permitted range, they are gone.
/// </summary>
/// <remarks>
/// This is a ref return, so if the underlaying value
/// changes, the returned value also changes. To avoid,
/// just don't use <c>ref</c> and structs will be copied.
/// </remarks>
/// <typeparam name="TElement">The type of data stored in the slice.</typeparam>
internal partial class Slice<TElement> : ISlice
{
    private static readonly Lut<TElement> Dummy = new();
    private readonly TElement _defaultValue = default;

    /// <summary>
    /// The content of the slice. Note that LUT uses index that starts from 0,
    /// so rows and columns must be adjusted to retrieved the value.
    /// </summary>
    private readonly Lut<Lut<TElement>> _data;

    /// <summary>
    /// Key is column number, value is number of cells in the column that are used.
    /// </summary>
    private readonly Dictionary<int, int> _columnUsage = new();

    internal Slice() => this._data = new Lut<Lut<TElement>>();

    /// <summary>
    /// Get the slice value at the specified point of the sheet.
    /// </summary>
    internal ref readonly TElement this[Point point] => ref this[point.Row, point.Column];

    /// <summary>
    /// Get the slice value at the specified point of the sheet.
    /// </summary>
    internal ref readonly TElement this[int row, int column]
    {
        get
        {
            Lut<TElement> rowLut = this._data.Get(row - 1);
            if (rowLut is null)
            {
                return ref this._defaultValue;
            }

            return ref rowLut.Get(column - 1);
        }
    }

    /// <inheritdoc />
    public bool IsEmpty => this.MaxRow == 0;

    /// <inheritdoc />
    public int MaxColumn { get; private set; }

    /// <inheritdoc />
    public int MaxRow => this._data.MaxUsedIndex + 1;

    /// <inheritdoc />
    public IEnumerable<int> UsedRows
    {
        get
        {
            Lut<Lut<TElement>>.LutEnumerator rowsEnumerator = new(
                this._data,
                XlsxSharp.XLHelper.MinRowNumber - 1,
                XlsxSharp.XLHelper.MaxRowNumber - 1
            );
            while (rowsEnumerator.MoveNext())
            {
                if (!rowsEnumerator.Current.IsEmpty)
                {
                    yield return rowsEnumerator.Index + 1;
                }
            }
        }
    }

    /// <inheritdoc />
    public Dictionary<int, int>.KeyCollection UsedColumns => this._columnUsage.Keys;

    /// <inheritdoc />
    public void Clear(Area area)
    {
        Enumerator enumerator = new(this, area);
        while (enumerator.MoveNext())
        {
            this.Set(enumerator.Point, in this._defaultValue);
        }
    }

    /// <inheritdoc />
    public void DeleteAreaAndShiftLeft(Area areaToDelete)
    {
        this.Clear(areaToDelete);

        bool noCellsToShift = areaToDelete.LastPoint.Column == XlsxSharp.XLHelper.MaxColumnNumber;
        if (noCellsToShift)
        {
            return;
        }

        int shiftDistance = areaToDelete.Width;
        Area shiftRange = areaToDelete.RightRange();
        Enumerator cellEnumerator = new(this, shiftRange);
        while (cellEnumerator.MoveNext())
        {
            Point srcPoint = cellEnumerator.Point;
            Point dstPoint = new(srcPoint.Row, srcPoint.Column - shiftDistance);
            this.Set(dstPoint, in cellEnumerator.Current);
            this.Set(srcPoint, in this._defaultValue);
        }
    }

    /// <inheritdoc />
    public void DeleteAreaAndShiftUp(Area areaToDelete)
    {
        this.Clear(areaToDelete);

        bool noCellsToShift = areaToDelete.LastPoint.Row == XlsxSharp.XLHelper.MaxRowNumber;
        if (noCellsToShift)
        {
            return;
        }

        int shiftDistance = areaToDelete.Height;
        Area shiftRange = areaToDelete.BelowRange();

        // Fast path for deleting full rows
        if (areaToDelete.HasFullRowWidth)
        {
            // Shifting full rows up to an empty space doesn't change column usage or max column and
            // is thus safe to only move row lookup tables to the new position. Start from top to not
            // overwrite not-yet moved rows.
            Lut<Lut<TElement>>.LutEnumerator rowEnumerator = new(
                this._data,
                shiftRange.TopRow - 1,
                shiftRange.BottomRow - 1
            );
            while (rowEnumerator.MoveNext())
            {
                // Enumerator is essentially a wrapped index and MoveNext() looks for the next
                // used index from current state of LUT, so it's fine to set the values like this.
                this._data.Set(rowEnumerator.Index - shiftDistance, rowEnumerator.Current);
                this._data.Set(rowEnumerator.Index, null);
            }

            return;
        }

        Enumerator cellEnumerator = new(this, shiftRange);
        while (cellEnumerator.MoveNext())
        {
            Point srcPoint = cellEnumerator.Point;
            Point dstPoint = new(srcPoint.Row - shiftDistance, srcPoint.Column);
            this.Set(dstPoint, in cellEnumerator.Current);
            this.Set(srcPoint, in this._defaultValue);
        }
    }

    /// <summary>
    /// Get enumerator over used values of the range.
    /// </summary>
    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false) =>
        !reverse ? new Enumerator(this, area) : new ReverseEnumerator(this, area);

    /// <inheritdoc />
    public void InsertAreaAndShiftDown(Area areaToInsert)
    {
        bool hasSpaceBelow = areaToInsert.LastPoint.Row < XlsxSharp.XLHelper.MaxRowNumber;
        if (!hasSpaceBelow)
        {
            this.Clear(areaToInsert);
            return;
        }

        int shiftDistance = areaToInsert.Height;

        // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
        Area purgedRange = new(
            new Point(
                XlsxSharp.XLHelper.MaxRowNumber - shiftDistance + 1,
                areaToInsert.FirstPoint.Column
            ),
            new Point(XlsxSharp.XLHelper.MaxRowNumber, areaToInsert.LastPoint.Column)
        );
        this.Clear(purgedRange);

        Area shiftedRange = new(
            areaToInsert.FirstPoint,
            new Point(
                XlsxSharp.XLHelper.MaxRowNumber - shiftDistance,
                areaToInsert.LastPoint.Column
            )
        );
        ReverseEnumerator cellEnumerator = new(this, shiftedRange);
        while (cellEnumerator.MoveNext())
        {
            Point srcPoint = cellEnumerator.Point;
            Point dstPoint = new(srcPoint.Row + shiftDistance, srcPoint.Column);
            this.Set(dstPoint, in cellEnumerator.Current);
            this.Set(srcPoint, in this._defaultValue);
        }
    }

    /// <inheritdoc />
    public void InsertAreaAndShiftRight(Area areaToInsert)
    {
        bool hasSpaceRight = areaToInsert.LastPoint.Column < XlsxSharp.XLHelper.MaxColumnNumber;
        if (!hasSpaceRight)
        {
            this.Clear(areaToInsert);
            return;
        }

        int shiftDistance = areaToInsert.Width;

        // Purged range might contain some cells that wouldn't be overwritten during shift => clear.
        Area purgedRange = new(
            new Point(
                areaToInsert.FirstPoint.Row,
                XlsxSharp.XLHelper.MaxColumnNumber - shiftDistance + 1
            ),
            new Point(areaToInsert.LastPoint.Row, XlsxSharp.XLHelper.MaxColumnNumber)
        );
        this.Clear(purgedRange);

        Area shiftedRange = new(
            areaToInsert.FirstPoint,
            new Point(
                areaToInsert.LastPoint.Row,
                XlsxSharp.XLHelper.MaxColumnNumber - shiftDistance
            )
        );
        ReverseEnumerator enumerator = new(this, shiftedRange);
        while (enumerator.MoveNext())
        {
            Point srcPoint = enumerator.Point;
            Point dstPoint = new(srcPoint.Row, srcPoint.Column + shiftDistance);
            this.Set(dstPoint, in enumerator.Current);
            this.Set(srcPoint, in this._defaultValue);
        }
    }

    public bool IsUsed(Point address)
    {
        Lut<TElement> rowLut = this._data.Get(address.Row - 1);
        if (rowLut is null)
        {
            return false;
        }

        return rowLut.IsUsed(address.Column - 1);
    }

    public void Swap(Point sp1, Point sp2)
    {
        TElement value1 = this[sp1];
        TElement value2 = this[sp2];
        this.Set(sp1, in value2);
        this.Set(sp2, in value1);
    }

    internal void SetAll(Area range, in TElement value)
    {
        foreach (Point point in range)
        {
            this.Set(point, in value);
        }
    }

    internal void Set(Point point, in TElement value) =>
        this.Set(point.Row, point.Column, in value);

    internal void Set(int row, int column, in TElement value)
    {
        Lut<TElement> rowLut = this._data.Get(row - 1);
        if (rowLut is null)
        {
            rowLut = new Lut<TElement>();
            this._data.Set(row - 1, rowLut);
        }

        bool wasUsed = rowLut.IsUsed(column - 1);
        rowLut.Set(column - 1, value);
        bool isUsed = rowLut.IsUsed(column - 1);

        if (wasUsed && !isUsed)
        {
            int newCount = this.DecrementColumnUsage(column);
            if (newCount == 0 && this.MaxColumn == column)
            {
                this.MaxColumn = this.CalculateMaxColumn();
            }

            if (rowLut.IsEmpty)
            {
                this._data.Set(row - 1, null);
            }
        }

        if (!wasUsed && isUsed)
        {
            this.IncrementColumnUsage(column);
            if (column > this.MaxColumn)
            {
                this.MaxColumn = column;
            }
        }
    }

    private int CalculateMaxColumn()
    {
        int maxColIdx = -1;
        Lut<Lut<TElement>>.LutEnumerator rowEnumerator = new(
            this._data,
            XlsxSharp.XLHelper.MinRowNumber - 1,
            XlsxSharp.XLHelper.MaxRowNumber - 1
        );
        while (rowEnumerator.MoveNext())
        {
            maxColIdx = Math.Max(maxColIdx, rowEnumerator.Current.MaxUsedIndex);
        }

        return maxColIdx + 1;
    }

    private int DecrementColumnUsage(int column)
    {
        if (!this._columnUsage.TryGetValue(column, out int count))
        {
            return 0;
        }

        if (count > 1)
        {
            return this._columnUsage[column] = count - 1;
        }

        this._columnUsage.Remove(column);
        return 0;
    }

    private void IncrementColumnUsage(int column)
    {
        if (this._columnUsage.TryGetValue(column, out int value))
        {
            this._columnUsage[column] = value + 1;
        }
        else
        {
            this._columnUsage.Add(column, 1);
        }
    }

    /// <summary>
    /// Enumerator that returns used values from a specified range.
    /// </summary>
    [DebuggerDisplay("{Point}:{Current}")]
    internal class Enumerator : IEnumerator<Point>
    {
        private readonly Area _range;
        private Lut<TElement>.LutEnumerator _columnsEnumerator;
        private Lut<Lut<TElement>>.LutEnumerator _rowsEnumerator;

        internal Enumerator(Slice<TElement> slice, Area range)
        {
            this._range = range;

            this._columnsEnumerator = new Lut<TElement>.LutEnumerator(
                Dummy,
                XlsxSharp.XLHelper.MaxColumnNumber + 1,
                XlsxSharp.XLHelper.MaxColumnNumber + 1
            );
            this._rowsEnumerator = new Lut<Lut<TElement>>.LutEnumerator(
                slice._data,
                range.FirstPoint.Row - 1,
                range.LastPoint.Row - 1
            );
        }

        public ref readonly TElement Current => ref this._columnsEnumerator.Current;

        public Point Point =>
            new(this._rowsEnumerator.Index + 1, this._columnsEnumerator.Index + 1);

        /// <summary>
        /// The movement is columns first, then rows.
        /// </summary>
        public bool MoveNext()
        {
            while (!this._columnsEnumerator.MoveNext())
            {
                if (!this._rowsEnumerator.MoveNext())
                {
                    return false;
                }

                this._columnsEnumerator = new Lut<TElement>.LutEnumerator(
                    this._rowsEnumerator.Current,
                    this._range.FirstPoint.Column - 1,
                    this._range.LastPoint.Column - 1
                );
            }

            return true;
        }

        void IEnumerator.Reset() => throw new NotSupportedException();

        Point IEnumerator<Point>.Current => this.Point;

        object IEnumerator.Current => this.Point;

        void IDisposable.Dispose() { }
    }

    [DebuggerDisplay("{Point}:{Current}")]
    private class ReverseEnumerator : IEnumerator<Point>
    {
        private readonly Area _range;
        private Lut<TElement>.ReverseLutEnumerator _columnsEnumerator;
        private Lut<Lut<TElement>>.ReverseLutEnumerator _rowsEnumerator;

        internal ReverseEnumerator(Slice<TElement> slice, Area range)
        {
            this._range = range;
            this._columnsEnumerator = new Lut<TElement>.ReverseLutEnumerator(Dummy, -1, -1);
            this._rowsEnumerator = new Lut<Lut<TElement>>.ReverseLutEnumerator(
                slice._data,
                range.FirstPoint.Row - 1,
                range.LastPoint.Row - 1
            );
        }

        public ref TElement Current => ref this._columnsEnumerator.Current;

        public Point Point =>
            new(this._rowsEnumerator.Index + 1, this._columnsEnumerator.Index + 1);

        public bool MoveNext()
        {
            while (!this._columnsEnumerator.MoveNext())
            {
                if (!this._rowsEnumerator.MoveNext())
                {
                    return false;
                }

                this._columnsEnumerator = new Lut<TElement>.ReverseLutEnumerator(
                    this._rowsEnumerator.Current,
                    this._range.FirstPoint.Column - 1,
                    this._range.LastPoint.Column - 1
                );
            }
            return true;
        }

        void IEnumerator.Reset() => throw new NotSupportedException();

        Point IEnumerator<Point>.Current => this.Point;

        object IEnumerator.Current => this.Point;

        public void Dispose() { }
    }
}

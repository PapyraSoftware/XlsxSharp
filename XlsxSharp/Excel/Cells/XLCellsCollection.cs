using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel;

internal class XLCellsCollection : IWorkbookListener
{
    private readonly XLWorksheet _ws;
    private readonly List<ISlice> _slices;

    public XLCellsCollection(XLWorksheet ws)
    {
        this._ws = ws;
        this.ValueSlice = new ValueSlice(ws.Workbook.SharedStringTable);
        this.FormulaSlice = new FormulaSlice(ws);
        this.FormatSlice = new FormatSlice();
        this._slices = [this.ValueSlice, this.FormulaSlice, this.FormatSlice, this.MiscSlice];
    }

    internal HashSet<int> ColumnsUsedKeys
    {
        get
        {
            HashSet<int> set = [];
            foreach (ISlice slice in this._slices)
            {
                set.UnionWith(slice.UsedColumns);
            }

            return set;
        }
    }

    internal bool IsEmpty => this._slices.All(slice => slice.IsEmpty);

    internal int MaxColumnUsed
    {
        get
        {
            int max = int.MinValue;
            foreach (ISlice slice in this._slices)
            {
                max = Math.Max(max, slice.MaxColumn);
            }

            return Math.Max(1, max);
        }
    }

    internal int MaxRowUsed
    {
        get
        {
            int max = int.MinValue;
            foreach (ISlice slice in this._slices)
            {
                max = Math.Max(max, slice.MaxRow);
            }

            return Math.Max(1, max);
        }
    }

    internal HashSet<int> RowsUsedKeys
    {
        get
        {
            HashSet<int> set = [];
            foreach (ISlice slice in this._slices)
            {
                set.UnionWith(slice.UsedRows);
            }

            return set;
        }
    }

    internal ValueSlice ValueSlice { get; }

    internal FormulaSlice FormulaSlice { get; }

    internal FormatSlice FormatSlice { get; }

    internal Slice<XLMiscSliceContent> MiscSlice { get; } = new();

    internal XLWorksheet Worksheet => this._ws;

    internal void Clear() => this.Clear(Area.Full);

    internal void Clear(Area clearRange)
    {
        foreach (ISlice slice in this._slices)
        {
            slice.Clear(clearRange);
        }
    }

    internal void DeleteAreaAndShiftLeft(Area rangeToDelete)
    {
        foreach (ISlice slice in this._slices)
        {
            slice.DeleteAreaAndShiftLeft(rangeToDelete);
        }
    }

    internal void DeleteAreaAndShiftUp(Area rangeToDelete)
    {
        foreach (ISlice slice in this._slices)
        {
            slice.DeleteAreaAndShiftUp(rangeToDelete);
        }
    }

    internal XLCell GetCell(Point address) => new(this._ws, address);

    /// <summary>
    /// Get all used cells in the worksheet.
    /// </summary>
    internal IEnumerable<XLCell> GetCells() => this.GetCells(Area.Full);

    /// <summary>
    /// Get all used cells in the worksheet that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(Func<XLCell, bool> predicate) =>
        this.GetCells(Area.Full, predicate);

    /// <summary>
    /// Get all used cells in the range that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(
        int rowStart,
        int columnStart,
        int rowEnd,
        int columnEnd,
        Func<XLCell, bool>? predicate = null
    ) => this.GetCells(new Area(rowStart, columnStart, rowEnd, columnEnd), predicate);

    /// <summary>
    /// Get all used cells in the range that satisfy the predicate.
    /// </summary>
    internal IEnumerable<XLCell> GetCells(Area range, Func<XLCell, bool>? predicate = null)
    {
        SlicesEnumerator enumerator = new(range, this);

        while (enumerator.MoveNext())
        {
            Point cellAddress = enumerator.Current;
            XLCell cell = this.GetCell(cellAddress);
            if (predicate == null || predicate(cell))
            {
                yield return cell;
            }
        }
    }

    internal IEnumerable<XLCell> GetCellsInColumn(int column) =>
        this.GetCells(1, column, XlsxSharp.XLHelper.MaxRowNumber, column);

    internal IEnumerable<XLCell> GetCellsInRow(int row) =>
        this.GetCells(row, 1, row, XlsxSharp.XLHelper.MaxColumnNumber);

    /// <summary>
    /// Get cell or null, if cell is not used.
    /// </summary>
    internal XLCell? GetUsedCell(Point address)
    {
        if (!this.IsUsed(address))
        {
            return null;
        }

        return this.GetCell(address);
    }

    internal int FirstColumnUsed(
        Area searchRange,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate = null
    ) => this.FindUsedColumn(searchRange, options, predicate, false);

    internal int FirstRowUsed(
        Area searchRange,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate = null
    ) => this.FindUsedRow(searchRange, options, predicate, false);

    internal void InsertAreaAndShiftDown(Area insertedRange)
    {
        foreach (ISlice slice in this._slices)
        {
            slice.InsertAreaAndShiftDown(insertedRange);
        }
    }

    internal void InsertAreaAndShiftRight(Area insertedRange)
    {
        foreach (ISlice slice in this._slices)
        {
            slice.InsertAreaAndShiftRight(insertedRange);
        }
    }

    internal int LastColumnUsed(
        Area searchRange,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate = null
    ) => this.FindUsedColumn(searchRange, options, predicate, true);

    internal int LastRowUsed(
        Area searchRange,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate = null
    ) => this.FindUsedRow(searchRange, options, predicate, true);

    /// <summary>
    /// Remap rows of a range.
    /// </summary>
    /// <param name="map">A sorted map of rows. The values must be resorted row numbers from <paramref name="sheetRange"/>.</param>
    /// <param name="sheetRange">Sheet that should have its rows rearranged.</param>
    internal void RemapRows(IList<int> map, Area sheetRange)
    {
        RemapRanges(map, sheetRange.TopRow, SwapRows);

        void SwapRows(int prevRowNumber, int currentRowNumber)
        {
            Area prevRowRange = new(
                new Point(prevRowNumber, sheetRange.LeftColumn),
                new Point(prevRowNumber, sheetRange.RightColumn)
            );
            Area currentRowRange = new(
                new Point(currentRowNumber, sheetRange.LeftColumn),
                new Point(currentRowNumber, sheetRange.RightColumn)
            );
            this.SwapRanges(prevRowRange, currentRowRange);
        }
    }

    /// <summary>
    /// Remap columns of a range.
    /// </summary>
    /// <param name="map">A sorted map of columns. The values must be resorted columns numbers from <paramref name="sheetRange"/>.</param>
    /// <param name="sheetRange">Sheet that should have its columns rearranged.</param>
    internal void RemapColumns(IList<int> map, Area sheetRange)
    {
        RemapRanges(map, sheetRange.LeftColumn, SwapColumns);

        void SwapColumns(int prevColNumber, int currentColNumber)
        {
            Area prevRowRange = new(
                new Point(sheetRange.TopRow, prevColNumber),
                new Point(sheetRange.BottomRow, prevColNumber)
            );
            Area currentRowRange = new(
                new Point(sheetRange.TopRow, currentColNumber),
                new Point(sheetRange.BottomRow, currentColNumber)
            );
            this.SwapRanges(prevRowRange, currentRowRange);
        }
    }

    private static void RemapRanges(IList<int> map, int indexOffset, Action<int, int> swapData)
    {
        for (int i = 0; i < map.Count; ++i)
        {
            int axisNumber = i + indexOffset;
            int dataAxisNumber = map[i];
            if (axisNumber == dataAxisNumber)
            {
                continue;
            }

            // Current row doesn't contain data it should, so it is a part of a permutation
            // loop. Go over each item in a loop and
            // We need to replace
            int prevNumber = axisNumber;
            int currentNumber = dataAxisNumber;
            int startLoopNumber = prevNumber;
            do
            {
                // Current row number contains data that should be on the previous row number,
                // so swap them. That will fix another link in a loop (the previous one), but
                // will keep current inconsistent, but that will be fixed when loop completes.
                swapData(prevNumber, currentNumber);

                // Because previous row number is already fixed and will no longer be touched
                // during loop fix, mark it as a row that contains correct data.
                map[prevNumber - indexOffset] = prevNumber;

                prevNumber = currentNumber;
                currentNumber = map[currentNumber - indexOffset];
            } while (currentNumber != startLoopNumber);

            // Although we don't have to swap the last one (N count loop needs only N-1 swaps),
            // we have to mark the last row mapping for the last link (the one before start).
            map[prevNumber - indexOffset] = prevNumber;
        }
    }

    private void SwapRanges(Area sheetRange1, Area sheetRange2)
    {
        int rowCount = sheetRange1.LastPoint.Row - sheetRange1.FirstPoint.Row + 1;
        int columnCount = sheetRange1.LastPoint.Column - sheetRange1.FirstPoint.Column + 1;
        for (int row = 0; row < rowCount; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                Point sp1 = new(
                    sheetRange1.FirstPoint.Row + row,
                    sheetRange1.FirstPoint.Column + column
                );
                Point sp2 = new(
                    sheetRange2.FirstPoint.Row + row,
                    sheetRange2.FirstPoint.Column + column
                );

                this.SwapCellsContent(sp1, sp2);
            }
        }
    }

    private int FindUsedColumn(
        Area range,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate,
        bool descending
    )
    {
        IEnumerable<int> usedColumns = Enumerable.Empty<int>();
        foreach (ISlice slice in this._slices)
        {
            usedColumns = usedColumns.Concat(slice.UsedColumns);
        }

        usedColumns = usedColumns
            .Where(c => c >= range.FirstPoint.Column && c <= range.LastPoint.Column)
            .Distinct();
        usedColumns = descending
            ? usedColumns.OrderByDescending(x => x)
            : usedColumns.OrderBy(x => x);

        foreach (int columnNumber in usedColumns)
        {
            SlicesEnumerator enumerator = new(
                new Area(range.FirstPoint.Row, columnNumber, range.LastPoint.Row, columnNumber),
                this
            );
            while (enumerator.MoveNext())
            {
                XLCell cell = new(this._ws, enumerator.Current);
                if (!cell.IsEmpty(options) && (predicate == null || predicate(cell)))
                {
                    return enumerator.Current.Column;
                }
            }
        }

        return 0;
    }

    private int FindUsedRow(
        Area searchRange,
        XLCellsUsedOptions options,
        Func<IXLCell, bool>? predicate,
        bool reverse
    )
    {
        SlicesEnumerator enumerator = new(searchRange, this, reverse);

        while (enumerator.MoveNext())
        {
            Point cellAddress = enumerator.Current;
            XLCell cell = this.GetCell(cellAddress);
            if (!cell.IsEmpty(options) && (predicate == null || predicate(cell)))
            {
                return cellAddress.Row;
            }
        }

        return 0;
    }

    private bool IsUsed(Point address)
    {
        // This is different from XLCellUsedOptions, which uses a business logic (e.g. empty string is considered not-used).
        // Here, we ask whether any slice contains a used elements which might differ from cell used logic.
        foreach (ISlice slice in this._slices)
        {
            if (slice.IsUsed(address))
            {
                return true;
            }
        }

        return false;
    }

    internal void SwapCellsContent(Point sp1, Point sp2)
    {
        this.ValueSlice.Swap(sp1, sp2);
        this.FormulaSlice.Swap(sp1, sp2);
        this.FormatSlice.Swap(sp1, sp2);
        this.MiscSlice.Swap(sp1, sp2);
    }

    /// <summary>
    /// Gets used points in the range.
    /// </summary>
    internal SlicesEnumerator ForValuesAndFormulas(Area range)
    {
        IEnumerator<Point> valueEnumerator = this.ValueSlice.GetEnumerator(range);
        IEnumerator<Point> formulaEnumerator = this.FormulaSlice.GetEnumerator(range);
        return new SlicesEnumerator(false, valueEnumerator, formulaEnumerator);
    }

    /// <summary>
    /// Apply a deterministic format change on used cells.
    /// </summary>
    /// <remarks>
    /// Deterministic = when inputs are equal, outputs must also be equal. That is needed to
    /// cache format modifications. This method doesn't register formats in the workbook
    /// styles, it only sets values in the slice.
    /// </remarks>
    /// <param name="area">Area that is used to check for used cells.</param>
    /// <param name="modification">A deterministic modification. It should ensure the returned formats are registered in workbook styles.</param>
    /// <param name="resolver">A provider of format for non-materialized cells (e.g. column has a format and thus non-materialized cells should use column format).</param>
    internal void ApplyFormatOnUsed(
        Area area,
        Func<XLCellFormatValue, XLCellFormatValue> modification,
        Func<Point, XLCellFormatValue> resolver
    )
    {
        SlicesEnumerator enumerator = new(area, this);
        this.ApplyFormat(enumerator, modification, resolver);
    }

    /// <inheritdoc cref="ApplyFormatOnAll(Area, Func{XLCellFormatValue, XLCellFormatValue}, Func{Point, XLCellFormatValue})"/>
    /// <remarks>Unlike general purpose method, the modification function of this one doesn't require explicit
    /// registration of format into the <see cref="XLWorkbookStyles"/>.
    /// </remarks>
    /// <param name="area">Area that will have its format modified.</param>
    /// <param name="modififyBorder">Return a modified border of a format. Must be deterministic.</param>
    internal void ApplyFormatOnAll(
        Area area,
        Func<XLBorderFormatValue, XLBorderFormatValue> modififyBorder
    )
    {
        XLWorkbookStyles styles = this.Worksheet.Workbook.Styles;
        Func<XLCellFormatValue, XLCellFormatValue> modifyFormat = format =>
        {
            XLBorderFormatValue modifiedBorder = styles.GetRegisteredBorderFormat(
                format.Border,
                modififyBorder
            );
            XLCellFormatValue modifiedFormat = format with
            {
                Border = modifiedBorder,
                CustomFormat = format.CustomFormat | CellFormatComponents.Border,
            };
            return styles.GetRegisteredCellFormat(modifiedFormat, static f => f);
        };
        this.ApplyFormatOnAll(area, modifyFormat, point => this.Worksheet.GetStyleValue(point));
    }

    /// <summary>
    /// Apply a deterministic format change on all cells in <paramref name="area"/>.
    /// </summary>
    /// <inheritdoc cref="ApplyFormatOnUsed"/>
    internal void ApplyFormatOnAll(
        Area area,
        Func<XLCellFormatValue, XLCellFormatValue> modification,
        Func<Point, XLCellFormatValue> resolver
    )
    {
        using IEnumerator<Point> areaEnumerator = area.GetEnumerator();
        SlicesEnumerator enumerator = new(false, areaEnumerator);
        this.ApplyFormat(enumerator, modification, resolver);
    }

    private void ApplyFormat(
        SlicesEnumerator enumerator,
        Func<XLCellFormatValue, XLCellFormatValue> modification,
        Func<Point, XLCellFormatValue> resolver
    )
    {
        Dictionary<XLCellFormatValue, XLCellFormatValue> cache = new(
            ReferenceEqualityComparer<XLCellFormatValue>.Instance
        );
        while (enumerator.MoveNext())
        {
            Point point = enumerator.Current;
            XLCellFormatValue format = this.FormatSlice.GetFormat(point) ?? resolver(point);
            if (!cache.TryGetValue(format, out XLCellFormatValue? modifiedFormat))
            {
                modifiedFormat = modification(format);
                cache.Add(format, modifiedFormat);
            }

            this.FormatSlice.Set(point, modifiedFormat);
        }
    }

    /// <summary>
    /// Enumerator that combines several other slice enumerators and enumerates
    /// <see cref="Point"/> in any of them.
    /// </summary>
    internal struct SlicesEnumerator
    {
        private readonly List<IEnumerator<Point>> _enumerators;
        private readonly bool _reverse;

        public SlicesEnumerator(Area range, XLCellsCollection cellsCollection, bool reverse = false)
            : this(
                reverse,
                cellsCollection.ValueSlice.GetEnumerator(range, reverse),
                cellsCollection.FormulaSlice.GetEnumerator(range, reverse),
                cellsCollection.FormatSlice.GetEnumerator(range, reverse),
                cellsCollection.MiscSlice.GetEnumerator(range, reverse)
            ) { }

        public SlicesEnumerator(bool reverse, params IEnumerator<Point>[] enumerators)
        {
            this.Current = new Point(1, 1);
            this._reverse = reverse;
            this._enumerators = [];
            foreach (IEnumerator<Point> enumerator in enumerators)
            {
                if (enumerator.MoveNext())
                {
                    this._enumerators.Add(enumerator);
                }
            }
        }

        public Point Current { get; private set; }

        public bool MoveNext()
        {
            Point? current = null;
            for (int i = 0; i < this._enumerators.Count; ++i)
            {
                IEnumerator<Point> enumerator = this._enumerators[i];
                if (
                    current is null
                    || (
                        this._reverse
                            ? enumerator.Current.CompareTo(current.Value) > 0
                            : enumerator.Current.CompareTo(current.Value) < 0
                    )
                )
                {
                    current = enumerator.Current;
                }
            }

            if (current == null)
            {
                return false;
            }

            this.Current = current.Value;

            for (int i = this._enumerators.Count - 1; i >= 0; --i)
            {
                IEnumerator<Point> enumerator = this._enumerators[i];
                if (enumerator.Current == current)
                {
                    bool isDone = !enumerator.MoveNext();
                    if (isDone)
                    {
                        this._enumerators.RemoveAt(i);
                    }
                }
            }

            return true;
        }
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        using Slice<XLCellFormula>.Enumerator enumerator = this.FormulaSlice.GetForwardEnumerator(
            Area.Full
        );
        while (enumerator.MoveNext())
        {
            ref readonly XLCellFormula cellFormula = ref enumerator.Current;
            Point currentPoint = enumerator.Point;
            if (cellFormula.Type != FormulaType.Normal)
            {
                // Array or data formula. Only change name once, on master cell.
                bool isMasterCell = cellFormula.Range.FirstPoint == currentPoint;
                if (!isMasterCell)
                {
                    continue;
                }
            }

            cellFormula.RenameSheet(currentPoint, oldSheetName, newSheetName);
        }
    }
}

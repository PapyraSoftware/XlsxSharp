using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLCells : IXLCells, IEnumerable<XLCell>
{
    private readonly XLWorkbook _workbook;
    private readonly List<XLRangeAddress> _rangeAddresses = [];
    private readonly bool _usedCellsOnly;
    private readonly Func<XLCell, bool> _predicate;
    private readonly XLCellsUsedOptions _options;

    public XLCells(
        XLWorksheet worksheet,
        bool usedCellsOnly,
        XLCellsUsedOptions options,
        Func<XLCell, bool>? predicate = null
    )
        : this(worksheet.Workbook, usedCellsOnly, options, predicate) { }

    public XLCells(
        XLWorkbook workbook,
        bool usedCellsOnly,
        XLCellsUsedOptions options,
        Func<XLCell, bool>? predicate = null
    )
    {
        this._workbook = workbook;
        this._usedCellsOnly = usedCellsOnly;
        this._options = options;
        this._predicate = predicate ?? (_ => true);
    }

    #region IEnumerable<XLCell> Members

    private IEnumerable<XLCell> GetAllCells()
    {
        IEnumerable<IGrouping<XLWorksheet?, XLRangeAddress>> groupedAddresses =
            this._rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (IGrouping<XLWorksheet?, XLRangeAddress> worksheetGroup in groupedAddresses)
        {
            XLWorksheet ws = worksheetGroup.Key!;
            IEnumerable<Point> sheetPoints = worksheetGroup
                .SelectMany(addr => GetAllCellsInRange(addr))
                .Distinct();
            foreach (Point sheetPoint in sheetPoints)
            {
                XLCell c = ws.Cell(sheetPoint.Row, sheetPoint.Column);
                if (this._predicate(c))
                {
                    yield return c;
                }
            }
        }
    }

    private static IEnumerable<Point> GetAllCellsInRange(IXLRangeAddress rangeAddress)
    {
        if (!rangeAddress.IsValid)
        {
            yield break;
        }

        XLRangeAddress normalizedAddress = ((XLRangeAddress)rangeAddress).Normalize();
        int minRow = normalizedAddress.FirstAddress.RowNumber;
        int maxRow = normalizedAddress.LastAddress.RowNumber;
        int minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        int maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        for (int ro = minRow; ro <= maxRow; ro++)
        {
            for (int co = minColumn; co <= maxColumn; co++)
            {
                yield return new Point(ro, co);
            }
        }
    }

    private IEnumerable<XLCell> GetUsedCells()
    {
        HashSet<XLAddress> visitedCells = [];
        IEnumerable<IGrouping<XLWorksheet?, XLRangeAddress>> groupedAddresses =
            this._rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (IGrouping<XLWorksheet?, XLRangeAddress> worksheetGroup in groupedAddresses)
        {
            XLWorksheet ws = worksheetGroup.Key!;

            IEnumerable<Point> usedCellsCandidates = this.GetUsedCellsCandidates(ws);

            IOrderedEnumerable<XLCell> cells = worksheetGroup
                .SelectMany(addr => this.GetUsedCellsInRange(addr, ws, usedCellsCandidates))
                .OrderBy(cell => cell.Address.RowNumber)
                .ThenBy(cell => cell.Address.ColumnNumber);

            visitedCells.Clear();
            foreach (XLCell cell in cells)
            {
                if (visitedCells.Add(cell.Address))
                {
                    yield return cell;
                }
            }
        }
    }

    private IEnumerable<XLCell> GetUsedCellsInRange(
        XLRangeAddress rangeAddress,
        XLWorksheet worksheet,
        IEnumerable<Point> usedCellsCandidates
    )
    {
        if (!rangeAddress.IsValid)
        {
            yield break;
        }

        XLRangeAddress normalizedAddress = rangeAddress.Normalize();
        int minRow = normalizedAddress.FirstAddress.RowNumber;
        int maxRow = normalizedAddress.LastAddress.RowNumber;
        int minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        int maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        IEnumerable<XLCell> cellRange = worksheet.Internals.CellsCollection.GetCells(
            minRow,
            minColumn,
            maxRow,
            maxColumn,
            this._predicate
        );

        foreach (XLCell cell in cellRange)
        {
            if (!cell.IsEmpty(this._options) && this._predicate(cell))
            {
                yield return cell;
            }
        }

        foreach (Point sheetPoint in usedCellsCandidates)
        {
            if (
                sheetPoint.Row.Between(minRow, maxRow)
                && sheetPoint.Column.Between(minColumn, maxColumn)
            )
            {
                XLCell cell = worksheet.Cell(sheetPoint.Row, sheetPoint.Column);

                if (this._predicate(cell))
                {
                    yield return cell;
                }
            }
        }
    }

    private IEnumerable<Point> GetUsedCellsCandidates(XLWorksheet worksheet)
    {
        IEnumerable<Point> candidates = Enumerable.Empty<Point>();

        if (this._options == XLCellsUsedOptions.AllContents)
        {
            return candidates;
        }

        if (this._options.HasFlag(XLCellsUsedOptions.MergedRanges))
        {
            candidates = candidates.Union(
                worksheet.Internals.MergedRanges.SelectMany<XLRange, Point>(r =>
                    GetAllCellsInRange(r.RangeAddress)
                )
            );
        }

        if (this._options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
        {
            candidates = candidates.Union(
                worksheet.ConditionalFormats.SelectMany<XLConditionalFormat, Point>(cf =>
                    cf.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))
                )
            );
        }

        if (this._options.HasFlag(XLCellsUsedOptions.DataValidation))
        {
            candidates = candidates.Union(
                worksheet.DataValidations.SelectMany<XLDataValidation, Point>(dv =>
                    dv.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))
                )
            );
        }

        if (this._options.HasFlag(XLCellsUsedOptions.Sparklines))
        {
            candidates = candidates.Union(
                worksheet
                    .SparklineGroups.SelectMany(sg => sg)
                    .Select(sl => Point.FromAddress(sl.Location.Address))
            );
        }

        return candidates.Distinct();
    }

    public IEnumerator<XLCell> GetEnumerator() => this.GetCells().GetEnumerator();

    private IEnumerable<XLCell> GetCells() =>
        this._usedCellsOnly ? this.GetUsedCells() : this.GetAllCells();

    #endregion IEnumerable<XLCell> Members

    #region IXLCells Members

    IEnumerator<IXLCell> IEnumerable<IXLCell>.GetEnumerator() => this.GetCells().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public XLCellValue Value
    {
        set => this.ForEach<XLCell>(c => c.Value = value);
    }

    public IXLCells Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this.ForEach<XLCell>(c => c.Clear(clearOptions));
        return this;
    }

    public void DeleteComments() => this.ForEach<XLCell>(c => c.DeleteComment());

    public void DeleteSparklines() => this.ForEach<XLCell>(c => c.DeleteSparkline());

    public string FormulaA1
    {
        set => this.ForEach<XLCell>(c => c.FormulaA1 = value);
    }

    public string FormulaR1C1
    {
        set => this.ForEach<XLCell>(c => c.FormulaR1C1 = value);
    }

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    internal XLCellFormat Format
    {
        get
        {
            // For backwards compatibility, the sheet is considered the inner style. A terrible
            // choice, but it is what it is.
            XLWorksheet? sheet = this
                ._rangeAddresses.Select(x => x.Worksheet)
                .FirstOrDefault(x => x is not null);
            SheetArea[] areas = [.. this._rangeAddresses.Select(SheetArea.From)];
            return XLCellFormat.ForCells(this._workbook, areas, sheet);
        }
    }

    #endregion IXLCells Members

    public void Add(XLRangeAddress rangeAddress) => this._rangeAddresses.Add(rangeAddress);

    public void Add(XLCell cell) => this.Add(new XLRangeAddress(cell.Address, cell.Address));

    public void Select()
    {
        foreach (XLCell cell in this)
        {
            cell.Select();
        }
    }
}

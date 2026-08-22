#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal abstract class XLRangeBase : IXLRangeBase
{
    private XLSortElements _sortRows;
    private XLSortElements _sortColumns;

    protected XLRangeBase(XLRangeAddress rangeAddress) => this._rangeAddress = rangeAddress;

    /// <summary>
    /// Get format API object tailored to the range type.
    /// </summary>
    internal abstract XLCellFormat Format { get; }

    protected virtual void OnRangeAddressChanged(
        XLRangeAddress oldAddress,
        XLRangeAddress newAddress
    ) => this.Worksheet.RelocateRange(this.RangeType, oldAddress, newAddress);

    #region Public properties

    private XLRangeAddress _rangeAddress;

    public XLRangeAddress RangeAddress
    {
        get => this._rangeAddress;
        protected set
        {
            if (this._rangeAddress != value)
            {
                XLRangeAddress oldAddress = this._rangeAddress;
                this._rangeAddress = value;
                this.OnRangeAddressChanged(oldAddress, this._rangeAddress);
            }
        }
    }

    public XLWorksheet Worksheet => this.RangeAddress.Worksheet;

    internal Area SheetRange
    {
        get
        {
            if (!this.RangeAddress.IsValid)
            {
                throw new InvalidOperationException("Range address is invalid.");
            }

            return Area.FromRangeAddress(this.RangeAddress);
        }
    }

    IXLDataValidation IXLRangeBase.CreateDataValidation() => this.CreateDataValidation();

    internal XLDataValidation CreateDataValidation() =>
        this.Worksheet.DataValidations.Create(this.SheetRange);

    public IXLDataValidation GetDataValidation()
    {
        this.Worksheet.DataValidations.TryGet(
            this.RangeAddress,
            out IXLDataValidation existingDataValidation
        );
        return existingDataValidation;
    }

    #region IXLRangeBase Members

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    IXLRangeAddress IXLAddressable.RangeAddress => this.RangeAddress;

    IXLWorksheet IXLRangeBase.Worksheet => this.RangeAddress.Worksheet;

    public string FormulaA1
    {
        set =>
            this.Cells()
                .ForEach(c =>
                {
                    c.FormulaA1 = value;
                    c.FormulaReference = this.RangeAddress;
                });
    }

    public string FormulaArrayA1
    {
        set
        {
            Area range = Area.FromRangeAddress(this.RangeAddress);
            if (this.Worksheet.MergedRanges.Any(mr => mr.Intersects(this)))
            {
                throw new InvalidOperationException(
                    "Can't create array function over a merged range."
                );
            }

            if (this.Worksheet.Tables.Any<XLTable>(t => t.Intersects(this)))
            {
                throw new InvalidOperationException("Can't create array function over a table.");
            }

            if (
                this.Cells(false)
                    .Any<XLCell>(c =>
                        c.HasArrayFormula && !this.RangeAddress.ContainsWhole(c.FormulaReference)
                    )
            )
            {
                throw new InvalidOperationException(
                    "Can't create array function that partially covers another array function."
                );
            }

            string formula = value.TrimFormulaEqual();
            string fixedFunctionsFormula = FormulaTransformation.FixFutureFunctions(
                formula,
                this.Worksheet.Name,
                this.SheetRange.FirstPoint
            );
            XLCellFormula arrayFormula = XLCellFormula.Array(fixedFunctionsFormula, range, false);

            FormulaSlice formulaSlice = this.Worksheet.Internals.CellsCollection.FormulaSlice;
            formulaSlice.SetArray(range, arrayFormula);

            // If formula evaluates to a text, it is stored directly in a worksheet, not in SST. Thus
            // when the switch to formula happens, disable shared string and enable when formula is removed.
            ValueSlice valueSlice = this.Worksheet.Internals.CellsCollection.ValueSlice;
            for (int row = range.TopRow; row <= range.BottomRow; ++row)
            {
                for (int col = range.LeftColumn; col <= range.RightColumn; ++col)
                {
                    valueSlice.SetShareString(new Point(row, col), false);
                }
            }

            // Formula is shared across all cells, so it's enough to invalidate master cell
            XLCell masterCell = this.FirstCell();
            masterCell.InvalidateFormula();
        }
    }

    public string FormulaR1C1
    {
        set =>
            this.Cells()
                .ForEach(c =>
                {
                    c.FormulaR1C1 = value;
                    c.FormulaReference = this.RangeAddress;
                });
    }

    public bool ShareString
    {
        set => this.Cells().ForEach(c => c.ShareString = value);
    }

    public XLCellValue Value
    {
        set => this.Cells().ForEach(c => c.Value = value);
    }

    #endregion IXLRangeBase Members

    #endregion Public properties

    #region IXLRangeBase Members

    IXLCells IXLRangeBase.Cells(string cells) => this.Cells(cells);

    IXLCells IXLRangeBase.Cells(bool usedCellsOnly) => this.Cells(usedCellsOnly);

    IXLCells IXLRangeBase.Cells(bool usedCellsOnly, XLCellsUsedOptions options) =>
        this.Cells(usedCellsOnly, options);

    IXLCells IXLRangeBase.CellsUsed() => this.CellsUsed();

    IXLCell IXLRangeBase.FirstCell() => this.FirstCell();

    IXLCell IXLRangeBase.LastCell() => this.LastCell();

    IXLCell IXLRangeBase.FirstCellUsed() => this.FirstCellUsed(XLCellsUsedOptions.AllContents);

    IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options) => this.FirstCellUsed(options);

    IXLCell IXLRangeBase.FirstCellUsed(Func<IXLCell, bool> predicate) =>
        this.FirstCellUsed(predicate);

    IXLCell IXLRangeBase.FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate) =>
        this.FirstCellUsed(options, predicate);

    IXLCell IXLRangeBase.LastCellUsed() => this.LastCellUsed(XLCellsUsedOptions.AllContents);

    IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options) => this.LastCellUsed(options);

    IXLCell IXLRangeBase.LastCellUsed(Func<IXLCell, bool> predicate) =>
        this.LastCellUsed(predicate);

    IXLCell IXLRangeBase.LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate) =>
        this.LastCellUsed(options, predicate);

    public virtual IXLCells Cells() => this.Cells(false);

    public virtual XLCells Cells(bool usedCellsOnly) =>
        this.Cells(usedCellsOnly, XLCellsUsedOptions.AllContents);

    public XLCells Cells(bool usedCellsOnly, XLCellsUsedOptions options)
    {
        XLCells cells = new(this.Worksheet, usedCellsOnly, options) { this.RangeAddress };
        return cells;
    }

    public virtual XLCells Cells(string cells) => this.Ranges(cells).Cells();

    public IXLCells Cells(Func<IXLCell, bool> predicate)
    {
        XLCells cells = new(this.Worksheet, false, XLCellsUsedOptions.AllContents, predicate)
        {
            this.RangeAddress,
        };
        return cells;
    }

    public XLCells CellsUsed() => this.Cells(true);

    public IXLRange Merge() => this.Merge(true);

    public IXLRange Merge(bool checkIntersect)
    {
        if (this.RangeAddress.FirstAddress.Equals(this.RangeAddress.LastAddress))
        {
            return this.Worksheet.Range(this.RangeAddress);
        }

        XLRange asRange = this.AsRange();

        if (checkIntersect)
        {
            List<IXLRange> intersectedMergedRanges =
            [
                .. this.Worksheet.Internals.MergedRanges.GetIntersectedRanges(this.RangeAddress),
            ];
            foreach (IXLRange intersectedMergedRange in intersectedMergedRanges)
            {
                this.Worksheet.Internals.MergedRanges.Remove(intersectedMergedRange);
            }

            XLCell firstCell = this.FirstCell();
            List<XLCell> cellsUsed =
            [
                .. this.CellsUsedInternal(
                    XLCellsUsedOptions.All & ~XLCellsUsedOptions.MergedRanges,
                    c => c.Point != firstCell.Point
                ),
            ];
            cellsUsed.ForEach(c =>
                c.Clear(
                    XLClearOptions.All
                        & ~XLClearOptions.MergedRanges
                        & ~XLClearOptions.NormalFormats
                )
            );

            // When a range is merged, remaining cells of the area take on the format of the first cell
            if (firstCell.FormatValue is null)
            {
                this.Worksheet.Internals.CellsCollection.FormatSlice.Clear(this.SheetRange);
            }
            else
            {
                // Merging removes borders that are not consistent across the whole border, even on the first cell
                Area area = this.SheetRange;
                XLBorderLine? leftBorder = this.GetVerticalBorder(
                    area.LeftColumn,
                    area.TopRow,
                    area.BottomRow,
                    static b => b.Left
                );
                XLBorderLine? topBorder = this.GetHorizontalBorder(
                    area.TopRow,
                    area.LeftColumn,
                    area.RightColumn,
                    static b => b.Top
                );
                XLBorderLine? rightBorder = this.GetVerticalBorder(
                    area.RightColumn,
                    area.TopRow,
                    area.BottomRow,
                    static b => b.Right
                );
                XLBorderLine? bottomBorder = this.GetHorizontalBorder(
                    area.BottomRow,
                    area.LeftColumn,
                    area.RightColumn,
                    static b => b.Bottom
                );

                XLCellsCollection cellsCollection = this.Worksheet.Internals.CellsCollection;
                XLCellFormatValue borderlessFormat =
                    this.Worksheet.Workbook.Styles.GetModifiedFormat(
                        firstCell.FormatValue,
                        _ => XLBorderFormatValue.None
                    );
                cellsCollection.FormatSlice.SetAll(this.SheetRange, borderlessFormat);

                if (leftBorder is not null && leftBorder.Value.IsVisible)
                {
                    cellsCollection.ApplyFormatOnAll(
                        this.SheetRange.SliceFromLeft(1),
                        b => b with { Left = leftBorder.Value }
                    );
                }

                if (topBorder is not null && topBorder.Value.IsVisible)
                {
                    cellsCollection.ApplyFormatOnAll(
                        this.SheetRange.SliceFromTop(1),
                        b => b with { Top = topBorder.Value }
                    );
                }

                if (rightBorder is not null && rightBorder.Value.IsVisible)
                {
                    cellsCollection.ApplyFormatOnAll(
                        this.SheetRange.SliceFromRight(1),
                        b => b with { Right = rightBorder.Value }
                    );
                }

                if (bottomBorder is not null && bottomBorder.Value.IsVisible)
                {
                    cellsCollection.ApplyFormatOnAll(
                        this.SheetRange.SliceFromBottom(1),
                        b => b with { Bottom = bottomBorder.Value }
                    );
                }
            }
        }

        this.Worksheet.Internals.MergedRanges.Add(asRange);
        return asRange;
    }

    private XLBorderLine? GetHorizontalBorder(
        int row,
        int minColumn,
        int maxColumn,
        Func<XLBorderFormatValue, XLBorderLine> borderSide
    )
    {
        XLBorderLine initialSide = borderSide(
            this.Worksheet.GetStyleValue(new Point(row, minColumn)).Border
        );
        for (int column = minColumn + 1; column <= maxColumn; ++column)
        {
            XLBorderLine currentSide = borderSide(
                this.Worksheet.GetStyleValue(new Point(row, column)).Border
            );
            if (currentSide != initialSide)
            {
                return null;
            }
        }

        return initialSide;
    }

    private XLBorderLine? GetVerticalBorder(
        int column,
        int minRow,
        int maxRow,
        Func<XLBorderFormatValue, XLBorderLine> borderSide
    )
    {
        XLBorderLine initialSide = borderSide(
            this.Worksheet.GetStyleValue(new Point(minRow, column)).Border
        );
        for (int row = minRow + 1; row <= maxRow; ++row)
        {
            XLBorderLine currentSide = borderSide(
                this.Worksheet.GetStyleValue(new Point(row, column)).Border
            );
            if (currentSide != initialSide)
            {
                return null;
            }
        }

        return initialSide;
    }

    public IXLRange Unmerge()
    {
        string tAddress = this.RangeAddress.ToString();
        XLRange asRange = this.AsRange();
        if (
            this
                .Worksheet.Internals.MergedRanges.Select<XLRange, string>(m =>
                    m.RangeAddress.ToString()
                )
                .Any(mAddress => mAddress == tAddress)
        )
        {
            this.Worksheet.Internals.MergedRanges.Remove(asRange);
        }

        return asRange;
    }

    public IXLRangeBase Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        XLClearOptions cellClearOptions =
            clearOptions
            & ~XLClearOptions.ConditionalFormats
            & ~XLClearOptions.DataValidation
            & ~XLClearOptions.MergedRanges
            & ~XLClearOptions.Sparklines;
        XLCellsUsedOptions cellUsedOptions = cellClearOptions.ToCellsUsedOptions();
        foreach (XLCell cell in this.CellsUsedInternal(cellUsedOptions))
        {
            // We'll clear the conditional formatting, data validations, sparklines
            // and merged ranges later down.
            cell.Clear(cellClearOptions, true);
        }

        if (clearOptions.HasFlag(XLClearOptions.ConditionalFormats))
        {
            this.Worksheet.ConditionalFormats.Clear(this.SheetRange);
        }

        if (clearOptions.HasFlag(XLClearOptions.DataValidation))
        {
            XLDataValidation validation = this.CreateDataValidation();
            this.Worksheet.DataValidations.Delete(validation);
        }

        if (clearOptions.HasFlag(XLClearOptions.MergedRanges))
        {
            this.ClearMerged();
        }

        if (clearOptions.HasFlag(XLClearOptions.Sparklines))
        {
            this.RemoveSparklines();
        }

        if (clearOptions == XLClearOptions.All)
        {
            this.Worksheet.Internals.CellsCollection.Clear(
                Area.FromRangeAddress(this.RangeAddress)
            );
        }
        return this;
    }

    public IXLRangeBase Relative(IXLRangeBase sourceBaseRange, IXLRangeBase targetBaseRange)
    {
        XLRangeAddress xlSourceBaseRangeAddress = (XLRangeAddress)sourceBaseRange.RangeAddress;
        XLRangeAddress xlTargetBaseRangeAddress = (XLRangeAddress)targetBaseRange.RangeAddress;
        XLRangeAddress xlRangeAddress = this.RangeAddress.Relative(
            in xlSourceBaseRangeAddress,
            in xlTargetBaseRangeAddress
        );

        return ((XLRangeBase)targetBaseRange).Range(in xlRangeAddress);
    }

    internal void RemoveSparklines() =>
        this
            .Worksheet.SparklineGroups.GetSparklines(this)
            .ToList()
            .ForEach(sl => this.Worksheet.SparklineGroups.Remove(sl.Location));

    public void DeleteComments() => this.Cells().DeleteComments();

    public bool Contains(string rangeAddress)
    {
        string addressToUse = rangeAddress.Contains("!")
            ? rangeAddress.Substring(rangeAddress.IndexOf("!") + 1)
            : rangeAddress;

        XLAddress firstAddress;
        XLAddress lastAddress;
        if (addressToUse.Contains(':'))
        {
            string[] arrRange = addressToUse.Split(':');
            firstAddress = XLAddress.Create(this.Worksheet, arrRange[0]);
            lastAddress = XLAddress.Create(this.Worksheet, arrRange[1]);
        }
        else
        {
            firstAddress = XLAddress.Create(this.Worksheet, addressToUse);
            lastAddress = XLAddress.Create(this.Worksheet, addressToUse);
        }
        return this.Contains(firstAddress, lastAddress);
    }

    public bool Contains(IXLRangeBase range) =>
        this.Contains(
            (XLAddress)range.RangeAddress.FirstAddress,
            (XLAddress)range.RangeAddress.LastAddress
        );

    public bool Intersects(string rangeAddress) =>
        this.Intersects(this.Worksheet.Range(rangeAddress));

    public bool Intersects(IXLRangeBase range)
    {
        if (!range.RangeAddress.IsValid || !this.RangeAddress.IsValid)
        {
            return false;
        }

        IXLRangeAddress ma = range.RangeAddress;
        XLRangeAddress ra = this.RangeAddress;
        return ra.Intersects(ma);
    }

    IXLRange IXLRangeBase.AsRange() => this.AsRange();

    public virtual XLRange AsRange() => this.Worksheet.Range(this.RangeAddress);

    public IXLRange AddToNamed(string name) => this.AddToNamed(name, XLScope.Workbook);

    public IXLRange AddToNamed(string name, XLScope scope) => this.AddToNamed(name, scope, null);

    public IXLRange AddToNamed(string name, XLScope scope, string comment)
    {
        XLDefinedNames definedNames =
            scope == XLScope.Workbook
                ? this.Worksheet.Workbook.DefinedNamesInternal
                : this.Worksheet.DefinedNames;

        if (definedNames.TryGetScopedValue(name, out XLDefinedName definedName))
        {
            definedName.Add(this.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
        }
        else
        {
            definedNames.Add(
                name,
                this.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true),
                comment
            );
        }

        return this.AsRange();
    }

    public IXLRangeBase SetValue(XLCellValue value)
    {
        this.Cells().ForEach(c => c.SetValue(value));
        return this;
    }

    public bool IsMerged() => this.Cells().Any(c => c.IsMerged());

    public virtual bool IsEmpty() =>
        !this.CellsUsed().Any<XLCell>() || this.CellsUsed().Any<XLCell>(c => c.IsEmpty());

    public virtual bool IsEmpty(XLCellsUsedOptions options)
    {
        foreach (IXLCell cell in this.CellsUsed(options))
        {
            if (!cell.IsEmpty(options))
            {
                return false;
            }
        }
        return true;
    }

    public virtual bool IsEntireRow() => this.RangeAddress.IsEntireRow();

    public virtual bool IsEntireColumn() => this.RangeAddress.IsEntireColumn();

    public bool IsEntireSheet() => this.RangeAddress.IsEntireSheet();

    #endregion IXLRangeBase Members

    public IXLCells Search(
        string searchText,
        CompareOptions compareOptions = CompareOptions.Ordinal,
        bool searchFormulae = false
    )
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        return this.CellsUsed(
            XLCellsUsedOptions.AllContents,
            c =>
            {
                try
                {
                    if (searchFormulae)
                    {
                        return c.HasFormula
                                && culture.CompareInfo.IndexOf(
                                    c.FormulaA1,
                                    searchText,
                                    compareOptions
                                ) >= 0
                            || culture.CompareInfo.IndexOf(
                                c.Value.ToString(CultureInfo.CurrentCulture),
                                searchText,
                                compareOptions
                            ) >= 0;
                    }
                    else
                    {
                        return culture.CompareInfo.IndexOf(
                                c.Value.ToString(CultureInfo.CurrentCulture),
                                searchText,
                                compareOptions
                            ) >= 0;
                    }
                }
                catch
                {
                    return false;
                }
            }
        );
    }

    internal XLCell FirstCell() => this.Cell(1, 1);

    internal XLCell LastCell() => this.Cell(this.RowCount(), this.ColumnCount());

    internal XLCell FirstCellUsed() =>
        this.FirstCellUsed(XLCellsUsedOptions.AllContents, predicate: null);

    internal XLCell FirstCellUsed(Func<IXLCell, bool> predicate) =>
        this.FirstCellUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLCell FirstCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate = null)
    {
        List<IXLCell> cellsUsed =
        [
            .. this.CellsUsedInternal(options, r => r.FirstCell(), predicate),
        ];

        if (!cellsUsed.Any())
        {
            return null;
        }

        int firstRow = cellsUsed.Min(c => c.Address.RowNumber);
        int firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

        if (firstRow < this.RangeAddress.FirstAddress.RowNumber)
        {
            firstRow = this.RangeAddress.FirstAddress.RowNumber;
        }

        if (firstColumn < this.RangeAddress.FirstAddress.ColumnNumber)
        {
            firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
        }

        return this.Worksheet.Cell(firstRow, firstColumn);
    }

    internal XLCell LastCellUsed() =>
        this.LastCellUsed(XLCellsUsedOptions.AllContents, predicate: null);

    internal XLCell LastCellUsed(Func<IXLCell, bool> predicate) =>
        this.LastCellUsed(XLCellsUsedOptions.AllContents, predicate);

    internal XLCell LastCellUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate = null)
    {
        List<IXLCell> cellsUsed =
        [
            .. this.CellsUsedInternal(options, r => r.LastCell(), predicate),
        ];

        if (!cellsUsed.Any())
        {
            return null;
        }

        int lastRow = cellsUsed.Max(c => c.Address.RowNumber);
        int lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

        if (lastRow > this.RangeAddress.LastAddress.RowNumber)
        {
            lastRow = this.RangeAddress.LastAddress.RowNumber;
        }

        if (lastColumn > this.RangeAddress.LastAddress.ColumnNumber)
        {
            lastColumn = this.RangeAddress.LastAddress.ColumnNumber;
        }

        return this.Worksheet.Cell(lastRow, lastColumn);
    }

    internal XLCell Cell(int row, int column) => this.Cell(new Point(row, column));

    internal XLCell Cell(Point point) =>
        this.Cell(new XLAddress(this.Worksheet, point.Row, point.Column, false, false));

    public virtual XLCell Cell(string cellAddressInRange)
    {
        if (XlsxSharp.XLHelper.IsValidA1Address(cellAddressInRange))
        {
            return this.Cell(XLAddress.Create(this.Worksheet, cellAddressInRange));
        }

        if (
            this.Worksheet.DefinedNames.TryGetValue(
                cellAddressInRange,
                out IXLDefinedName definedName
            )
        )
        {
            return definedName.Ranges.First().FirstCell().CastTo<XLCell>();
        }

        return null;
    }

    public XLCell Cell(int row, string column) =>
        this.Cell(new XLAddress(this.Worksheet, row, column, false, false));

    public XLCell Cell(IXLAddress cellAddressInRange) =>
        this.Cell(cellAddressInRange.RowNumber, cellAddressInRange.ColumnNumber);

    public XLCell Cell(in XLAddress cellAddressInRange)
    {
        int absRow = cellAddressInRange.RowNumber + this.RangeAddress.FirstAddress.RowNumber - 1;
        int absColumn =
            cellAddressInRange.ColumnNumber + this.RangeAddress.FirstAddress.ColumnNumber - 1;

        if (absRow <= 0 || absRow > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellAddressInRange),
                string.Format(
                    "Row number must be between 1 and {0}",
                    XlsxSharp.XLHelper.MaxRowNumber
                )
            );
        }

        if (absColumn <= 0 || absColumn > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellAddressInRange),
                string.Format(
                    "Column number must be between 1 and {0}",
                    XlsxSharp.XLHelper.MaxColumnNumber
                )
            );
        }

        XLCell cell = this.Worksheet.Internals.CellsCollection.GetCell(
            new Point(absRow, absColumn)
        );
        return cell;
    }

    public int RowCount() =>
        this.RangeAddress.LastAddress.RowNumber - this.RangeAddress.FirstAddress.RowNumber + 1;

    public int RowCount(XLCellsUsedOptions cellsUsedOptions)
    {
        XLCell lcu = this.LastCellUsed(cellsUsedOptions);
        if (lcu == null)
        {
            return 0;
        }

        XLCell fcu = this.FirstCellUsed(cellsUsedOptions);
        if (fcu == null)
        {
            return 0;
        }

        return lcu.Address.RowNumber - fcu.Address.RowNumber + 1;
    }

    public int RowNumber() => this.RangeAddress.FirstAddress.RowNumber;

    public int ColumnCount() =>
        this.RangeAddress.LastAddress.ColumnNumber
        - this.RangeAddress.FirstAddress.ColumnNumber
        + 1;

    public int ColumnCount(XLCellsUsedOptions cellsUsedOptions)
    {
        XLCell lcu = this.LastCellUsed(cellsUsedOptions);
        if (lcu == null)
        {
            return 0;
        }

        XLCell fcu = this.FirstCellUsed(cellsUsedOptions);
        if (fcu == null)
        {
            return 0;
        }

        return lcu.Address.ColumnNumber - fcu.Address.ColumnNumber + 1;
    }

    public int ColumnNumber() => this.RangeAddress.FirstAddress.ColumnNumber;

    public string ColumnLetter() => this.RangeAddress.FirstAddress.ColumnLetter;

    public virtual XLRange Range(string rangeAddressStr)
    {
        XLRangeAddress rangeAddress = new(this.Worksheet, rangeAddressStr);
        return this.Range(rangeAddress);
    }

    internal abstract void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted);

    internal abstract void WorksheetRangeShiftedRows(XLRange range, int rowsShifted);

    public abstract XLRangeType RangeType { get; }

    public XLRange Range(IXLCell firstCell, IXLCell lastCell)
    {
        XLAddress newFirstCellAddress = (XLAddress)firstCell.Address;
        XLAddress newLastCellAddress = (XLAddress)lastCell.Address;

        return this.GetRange(newFirstCellAddress, newLastCellAddress);
    }

    private XLRange GetRange(XLAddress newFirstCellAddress, XLAddress newLastCellAddress)
    {
        if (!this.Worksheet.Equals(newFirstCellAddress.Worksheet))
        {
            throw new ArgumentException(
                "The address refers to a different worksheet.",
                nameof(newFirstCellAddress)
            );
        }

        if (!this.Worksheet.Equals(newLastCellAddress.Worksheet))
        {
            throw new ArgumentException(
                "The address refers to a different worksheet.",
                nameof(newLastCellAddress)
            );
        }

        XLRangeAddress newRangeAddress = new(newFirstCellAddress, newLastCellAddress);
        if (
            newFirstCellAddress.RowNumber < this.RangeAddress.FirstAddress.RowNumber
            || newFirstCellAddress.RowNumber > this.RangeAddress.LastAddress.RowNumber
            || newLastCellAddress.RowNumber > this.RangeAddress.LastAddress.RowNumber
            || newFirstCellAddress.ColumnNumber < this.RangeAddress.FirstAddress.ColumnNumber
            || newFirstCellAddress.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber
            || newLastCellAddress.ColumnNumber > this.RangeAddress.LastAddress.ColumnNumber
        )
        {
            throw new ArgumentOutOfRangeException(
                string.Format(
                    "The cells {0} and {1} are outside the range '{2}'.",
                    newFirstCellAddress,
                    newLastCellAddress,
                    this.ToString()
                )
            );
        }

        if (newFirstCellAddress.Worksheet != null)
        {
            return newFirstCellAddress.Worksheet.GetOrCreateRange(newRangeAddress);
        }
        else if (this.Worksheet != null)
        {
            return this.Worksheet.GetOrCreateRange(newRangeAddress);
        }
        else
        {
            return new XLRange(newRangeAddress, this.Style);
        }
    }

    public XLRange Range(string firstCellAddress, string lastCellAddress)
    {
        XLRangeAddress rangeAddress = new(
            XLAddress.Create(this.Worksheet, firstCellAddress),
            XLAddress.Create(this.Worksheet, lastCellAddress)
        );
        return this.Range(rangeAddress);
    }

    internal XLRange Range(Area area) =>
        this.Range(area.TopRow, area.LeftColumn, area.BottomRow, area.RightColumn);

    internal XLRange Range(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    )
    {
        XLRangeAddress rangeAddress = new(
            new XLAddress(
                this.Worksheet,
                firstCellRow + this.RangeAddress.FirstAddress.RowNumber - 1,
                firstCellColumn + this.RangeAddress.FirstAddress.ColumnNumber - 1,
                fixedRow: false,
                fixedColumn: false
            ),
            new XLAddress(
                this.Worksheet,
                lastCellRow + this.RangeAddress.FirstAddress.RowNumber - 1,
                lastCellColumn + this.RangeAddress.FirstAddress.ColumnNumber - 1,
                fixedRow: false,
                fixedColumn: false
            )
        );
        return this.Range(rangeAddress);
    }

    public XLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
    {
        XLRangeAddress rangeAddress = new((XLAddress)firstCellAddress, (XLAddress)lastCellAddress);
        return this.Range(rangeAddress);
    }

    public XLRange Range(IXLRangeAddress rangeAddress)
    {
        XLRangeAddress xlRangeAddress = (XLRangeAddress)rangeAddress;
        return this.Range(in xlRangeAddress);
    }

    internal XLRange Range(in XLRangeAddress rangeAddress)
    {
        XLWorksheet ws =
            rangeAddress.FirstAddress.Worksheet
            ?? rangeAddress.LastAddress.Worksheet
            ?? this.Worksheet;

        XLAddress newFirstCellAddress = new(
            ws,
            rangeAddress.FirstAddress.RowNumber,
            rangeAddress.FirstAddress.ColumnNumber,
            rangeAddress.FirstAddress.FixedRow,
            rangeAddress.FirstAddress.FixedColumn
        );

        XLAddress newLastCellAddress = new(
            ws,
            rangeAddress.LastAddress.RowNumber,
            rangeAddress.LastAddress.ColumnNumber,
            rangeAddress.LastAddress.FixedRow,
            rangeAddress.LastAddress.FixedColumn
        );

        return this.GetRange(newFirstCellAddress, newLastCellAddress);
    }

    public virtual XLRanges Ranges(string ranges)
    {
        XLRanges retVal = new(this.Worksheet);
        string[] rangePairs = ranges.Split(',');
        foreach (string pair in rangePairs)
        {
            retVal.Add(this.Range(pair.Trim()));
        }

        return retVal;
    }

    public IXLRanges Ranges(params string[] ranges)
    {
        XLRanges retVal = new(this.Worksheet);
        foreach (string pair in ranges)
        {
            retVal.Add(this.Range(pair));
        }

        return retVal;
    }

    protected string FixColumnAddress(string address)
    {
        if (int.TryParse(address, out int rowNumber))
        {
            return this.RangeAddress.FirstAddress.ColumnLetter
                + (rowNumber + this.RangeAddress.FirstAddress.RowNumber - 1).ToInvariantString();
        }

        return address;
    }

    protected string FixRowAddress(string address)
    {
        if (int.TryParse(address, out int columnNumber))
        {
            return XlsxSharp.XLHelper.GetColumnLetterFromNumber(
                    columnNumber + this.RangeAddress.FirstAddress.ColumnNumber - 1
                ) + this.RangeAddress.FirstAddress.RowNumber.ToInvariantString();
        }

        return address;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this.Worksheet, true, options) { this.RangeAddress };
        return cells;
    }

    public IXLCells CellsUsed(Func<IXLCell, bool> predicate)
    {
        XLCells cells = new(this.Worksheet, true, XLCellsUsedOptions.AllContents, predicate)
        {
            this.RangeAddress,
        };
        return cells;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options, Func<IXLCell, bool> predicate) =>
        this.CellsUsedInternal(options, predicate);

    internal XLCells CellsUsedInternal(
        XLCellsUsedOptions options,
        Func<XLCell, bool> predicate = null
    )
    {
        XLCells cells = new(this.Worksheet, true, options, predicate) { this.RangeAddress };
        return cells;
    }

    public IXLRangeColumns InsertColumnsAfter(int numberOfColumns) =>
        this.InsertColumnsAfter(numberOfColumns, true);

    public IXLRangeColumns InsertColumnsAfter(int numberOfColumns, bool expandRange)
    {
        IXLRangeColumns retVal = this.InsertColumnsAfter(false, numberOfColumns);
        // Adjust the range
        if (expandRange)
        {
            this.RangeAddress = new XLRangeAddress(
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.FirstAddress.FixedRow,
                    this.RangeAddress.FirstAddress.FixedColumn
                ),
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber + numberOfColumns,
                    this.RangeAddress.LastAddress.FixedRow,
                    this.RangeAddress.LastAddress.FixedColumn
                )
            );
        }
        return retVal;
    }

    public IXLRangeColumns InsertColumnsAfter(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true
    ) => this.InsertColumnsAfterInternal(onlyUsedCells, numberOfColumns, formatFromLeft);

    public void InsertColumnsAfterVoid(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true
    ) =>
        this.InsertColumnsAfterInternal(
            onlyUsedCells,
            numberOfColumns,
            formatFromLeft,
            nullReturn: true
        );

    private IXLRangeColumns InsertColumnsAfterInternal(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true,
        bool nullReturn = false
    )
    {
        int columnCount = this.ColumnCount();
        int firstColumn = this.RangeAddress.FirstAddress.ColumnNumber + columnCount;
        if (firstColumn > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            firstColumn = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        int lastColumn = firstColumn + this.ColumnCount() - 1;
        if (lastColumn > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumn = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        int firstRow = this.RangeAddress.FirstAddress.RowNumber;
        int lastRow = firstRow + this.RowCount() - 1;
        if (lastRow > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRow = XlsxSharp.XLHelper.MaxRowNumber;
        }

        XLRange newRange = this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        return newRange.InsertColumnsBeforeInternal(
            onlyUsedCells,
            numberOfColumns,
            formatFromLeft,
            nullReturn
        );
    }

    public IXLRangeColumns InsertColumnsBefore(int numberOfColumns) =>
        this.InsertColumnsBefore(numberOfColumns, false);

    public IXLRangeColumns InsertColumnsBefore(int numberOfColumns, bool expandRange)
    {
        IXLRangeColumns retVal = this.InsertColumnsBefore(false, numberOfColumns);
        // Adjust the range
        if (expandRange)
        {
            this.RangeAddress = new XLRangeAddress(
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.FirstAddress.ColumnNumber - numberOfColumns,
                    this.RangeAddress.FirstAddress.FixedRow,
                    this.RangeAddress.FirstAddress.FixedColumn
                ),
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.FixedRow,
                    this.RangeAddress.LastAddress.FixedColumn
                )
            );
        }
        return retVal;
    }

    public IXLRangeColumns InsertColumnsBefore(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true
    ) => this.InsertColumnsBeforeInternal(onlyUsedCells, numberOfColumns, formatFromLeft);

    public void InsertColumnsBeforeVoid(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true
    ) =>
        this.InsertColumnsBeforeInternal(
            onlyUsedCells,
            numberOfColumns,
            formatFromLeft,
            nullReturn: true
        );

    private IXLRangeColumns InsertColumnsBeforeInternal(
        bool onlyUsedCells,
        int numberOfColumns,
        bool formatFromLeft = true,
        bool nullReturn = false
    )
    {
        if (numberOfColumns <= 0 || numberOfColumns > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(numberOfColumns),
                $"Number of columns to insert must be a positive number no more than {XlsxSharp.XLHelper.MaxColumnNumber}"
            );
        }

        foreach (XLWorksheet ws in this.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (
                XLCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null)
            )
            {
                cell.ShiftFormulaColumns(this.AsRange(), numberOfColumns);
            }
        }

        // Inserting and shifting of whole columns is rather inconsistent across the codebase. In some places, the columns collection
        // is shifted before this method is called and thus the we can't shift column properties again. In others, the code relies on
        // shifting in this method.
        if (!onlyUsedCells)
        {
            int lastColumn = this.Worksheet.Internals.CellsCollection.MaxColumnUsed;
            if (lastColumn > 0)
            {
                int firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
                for (int co = lastColumn; co >= firstColumn; co--)
                {
                    int newColumn = co + numberOfColumns;
                    if (this.IsEntireColumn())
                    {
                        this.Worksheet.Column(newColumn).Width = this.Worksheet.Column(co).Width;
                    }
                }
            }
        }

        Area insertedRange = new(
            Point.FromAddress(this.RangeAddress.FirstAddress),
            new Point(
                this.RangeAddress.LastAddress.RowNumber,
                this.RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1
            )
        );

        this.Worksheet.Internals.CellsCollection.InsertAreaAndShiftRight(insertedRange);

        int firstRowReturn = this.RangeAddress.FirstAddress.RowNumber;
        int lastRowReturn = this.RangeAddress.LastAddress.RowNumber;
        int firstColumnReturn = this.RangeAddress.FirstAddress.ColumnNumber;
        int lastColumnReturn = this.RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

        this.Worksheet.NotifyRangeShiftedColumns(this.AsRange(), numberOfColumns);

        XLRange rangeToReturn = this.Worksheet.Range(
            firstRowReturn,
            firstColumnReturn,
            lastRowReturn,
            lastColumnReturn
        );

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        XLCellsUsedOptions contentFlags =
            XLCellsUsedOptions.All
            & ~XLCellsUsedOptions.ConditionalFormats
            & ~XLCellsUsedOptions.DataValidation;

        if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
        {
            XLRangeColumn firstColumnUsed = rangeToReturn.FirstColumn();
            XLRangeColumn model = firstColumnUsed.ColumnLeft();
            IXLCell modelFirstRow = (model as IXLRangeBase).FirstCellUsed(contentFlags);
            IXLCell modelLastRow = (model as IXLRangeBase).LastCellUsed(contentFlags);
            if (modelLastRow != null)
            {
                int firstRoReturned =
                    modelFirstRow.Address.RowNumber - model.RangeAddress.FirstAddress.RowNumber + 1;
                int lastRoReturned =
                    modelLastRow.Address.RowNumber - model.RangeAddress.FirstAddress.RowNumber + 1;
                for (int ro = firstRoReturned; ro <= lastRoReturned; ro++)
                {
                    rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
                }
            }
        }
        else
        {
            XLRangeRow lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
            if (lastRoUsed != null)
            {
                int lastRoReturned = lastRoUsed.RowNumber();
                for (int ro = 1; ro <= lastRoReturned; ro++)
                {
                    IXLStyle styleToUse = this.Worksheet.Internals.RowsCollection.TryGetValue(
                        ro,
                        out XLRow row
                    )
                        ? row.Style
                        : this.Worksheet.Style;

                    rangeToReturn.Row(ro).Style = styleToUse;
                }
            }
        }

        if (nullReturn)
        {
            return null;
        }

        return rangeToReturn.Columns();
    }

    public IXLRangeRows InsertRowsBelow(int numberOfRows) =>
        this.InsertRowsBelow(numberOfRows, true);

    public IXLRangeRows InsertRowsBelow(int numberOfRows, bool expandRange)
    {
        IXLRangeRows retVal = this.InsertRowsBelow(false, numberOfRows);
        // Adjust the range
        if (expandRange)
        {
            this.RangeAddress = new XLRangeAddress(
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.FirstAddress.RowNumber,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.FirstAddress.FixedRow,
                    this.RangeAddress.FirstAddress.FixedColumn
                ),
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.LastAddress.RowNumber + numberOfRows,
                    this.RangeAddress.LastAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.FixedRow,
                    this.RangeAddress.LastAddress.FixedColumn
                )
            );
        }
        return retVal;
    }

    public IXLRangeRows InsertRowsBelow(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove = true
    ) =>
        this.InsertRowsBelowInternal(
            onlyUsedCells,
            numberOfRows,
            formatFromAbove,
            nullReturn: false
        );

    public void InsertRowsBelowVoid(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove = true
    ) =>
        this.InsertRowsBelowInternal(
            onlyUsedCells,
            numberOfRows,
            formatFromAbove,
            nullReturn: true
        );

    private IXLRangeRows InsertRowsBelowInternal(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove,
        bool nullReturn
    )
    {
        int rowCount = this.RowCount();
        int firstRow = this.RangeAddress.FirstAddress.RowNumber + rowCount;
        if (firstRow > XlsxSharp.XLHelper.MaxRowNumber)
        {
            firstRow = XlsxSharp.XLHelper.MaxRowNumber;
        }

        int lastRow = firstRow + this.RowCount() - 1;
        if (lastRow > XlsxSharp.XLHelper.MaxRowNumber)
        {
            lastRow = XlsxSharp.XLHelper.MaxRowNumber;
        }

        int firstColumn = this.RangeAddress.FirstAddress.ColumnNumber;
        int lastColumn = firstColumn + this.ColumnCount() - 1;
        if (lastColumn > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            lastColumn = XlsxSharp.XLHelper.MaxColumnNumber;
        }

        XLRange newRange = this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        return newRange.InsertRowsAboveInternal(
            onlyUsedCells,
            numberOfRows,
            formatFromAbove,
            nullReturn
        );
    }

    public IXLRangeRows InsertRowsAbove(int numberOfRows) =>
        this.InsertRowsAbove(numberOfRows, false);

    public IXLRangeRows InsertRowsAbove(int numberOfRows, bool expandRange)
    {
        IXLRangeRows retVal = this.InsertRowsAbove(false, numberOfRows);
        // Adjust the range
        if (expandRange)
        {
            this.RangeAddress = new XLRangeAddress(
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.FirstAddress.RowNumber - numberOfRows,
                    this.RangeAddress.FirstAddress.ColumnNumber,
                    this.RangeAddress.FirstAddress.FixedRow,
                    this.RangeAddress.FirstAddress.FixedColumn
                ),
                new XLAddress(
                    this.Worksheet,
                    this.RangeAddress.LastAddress.RowNumber,
                    this.RangeAddress.LastAddress.ColumnNumber,
                    this.RangeAddress.LastAddress.FixedRow,
                    this.RangeAddress.LastAddress.FixedColumn
                )
            );
        }
        return retVal;
    }

    public void InsertRowsAboveVoid(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove = true
    ) =>
        this.InsertRowsAboveInternal(
            onlyUsedCells,
            numberOfRows,
            formatFromAbove,
            nullReturn: true
        );

    public IXLRangeRows InsertRowsAbove(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove = true
    ) =>
        this.InsertRowsAboveInternal(
            onlyUsedCells,
            numberOfRows,
            formatFromAbove,
            nullReturn: false
        );

    private IXLRangeRows InsertRowsAboveInternal(
        bool onlyUsedCells,
        int numberOfRows,
        bool formatFromAbove,
        bool nullReturn
    )
    {
        if (numberOfRows <= 0 || numberOfRows > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(numberOfRows),
                $"Number of rows to insert must be a positive number no more than {XlsxSharp.XLHelper.MaxRowNumber}"
            );
        }

        XLRange asRange = this.AsRange();
        foreach (XLWorksheet ws in this.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (
                XLCell cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null)
            )
            {
                cell.ShiftFormulaRows(asRange, numberOfRows);
            }
        }

        if (!onlyUsedCells)
        {
            int lastRow = this.Worksheet.Internals.CellsCollection.MaxRowUsed;
            if (lastRow > 0)
            {
                int firstRow = this.RangeAddress.FirstAddress.RowNumber;
                for (int ro = lastRow; ro >= firstRow; ro--)
                {
                    int newRow = ro + numberOfRows;
                    if (this.IsEntireRow())
                    {
                        this.Worksheet.Row(newRow).Height = this.Worksheet.Row(ro).Height;
                    }
                }
            }
        }

        Area insertedRange = new(
            Point.FromAddress(this.RangeAddress.FirstAddress),
            new Point(
                this.RangeAddress.FirstAddress.RowNumber + numberOfRows - 1,
                this.RangeAddress.LastAddress.ColumnNumber
            )
        );
        this.Worksheet.Internals.CellsCollection.InsertAreaAndShiftDown(insertedRange);

        int firstRowReturn = this.RangeAddress.FirstAddress.RowNumber;
        int lastRowReturn = this.RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
        int firstColumnReturn = this.RangeAddress.FirstAddress.ColumnNumber;
        int lastColumnReturn = this.RangeAddress.LastAddress.ColumnNumber;

        this.Worksheet.NotifyRangeShiftedRows(this.AsRange(), numberOfRows);

        XLRange rangeToReturn = this.Worksheet.Range(
            firstRowReturn,
            firstColumnReturn,
            lastRowReturn,
            lastColumnReturn
        );

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        XLCellsUsedOptions contentFlags =
            XLCellsUsedOptions.All
            & ~XLCellsUsedOptions.ConditionalFormats
            & ~XLCellsUsedOptions.DataValidation;

        if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
        {
            XLRangeRow fr = rangeToReturn.FirstRow();
            XLRangeRow model = fr.RowAbove();
            IXLCell modelFirstColumn = (model as IXLRangeBase).FirstCellUsed(contentFlags);
            IXLCell modelLastColumn = (model as IXLRangeBase).LastCellUsed(contentFlags);
            if (modelFirstColumn != null && modelLastColumn != null)
            {
                int firstCoReturned =
                    modelFirstColumn.Address.ColumnNumber
                    - model.RangeAddress.FirstAddress.ColumnNumber
                    + 1;
                int lastCoReturned =
                    modelLastColumn.Address.ColumnNumber
                    - model.RangeAddress.FirstAddress.ColumnNumber
                    + 1;
                for (int co = firstCoReturned; co <= lastCoReturned; co++)
                {
                    rangeToReturn.Column(co).Style = model.Cell(co).Style;
                }
            }
        }
        else
        {
            XLRangeColumn lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
            if (lastCoUsed != null)
            {
                int lastCoReturned = lastCoUsed.ColumnNumber();
                for (int co = 1; co <= lastCoReturned; co++)
                {
                    IXLStyle styleToUse = this.Worksheet.Internals.ColumnsCollection.TryGetValue(
                        co,
                        out XLColumn column
                    )
                        ? column.Style
                        : this.Worksheet.Style;

                    rangeToReturn.Style = styleToUse;
                }
            }
        }

        // Skip calling .Rows() for performance reasons if required.
        if (nullReturn)
        {
            return null;
        }

        return rangeToReturn.Rows();
    }

    private void ClearMerged()
    {
        List<IXLRange> mergeToDelete =
        [
            .. this.Worksheet.Internals.MergedRanges.GetIntersectedRanges(this.RangeAddress),
        ];
        mergeToDelete.ForEach(m => this.Worksheet.Internals.MergedRanges.Remove(m));
    }

    public bool Contains(IXLCell cell) => this.Contains((XLAddress)cell.Address);

    public bool Contains(XLAddress first, XLAddress last) =>
        this.Contains(first) && this.Contains(last);

    public bool Contains(XLAddress address) => this.RangeAddress.Contains(in address);

    public void Delete(XLShiftDeletedCells shiftDeleteCells)
    {
        int numberOfRows = this.RowCount();
        int numberOfColumns = this.ColumnCount();

        if (!this.RangeAddress.IsValid)
        {
            return;
        }

        this.Worksheet.SparklineGroups.Remove(this);

        IXLRange shiftedRangeFormula = this.Worksheet.Range(
            this.RangeAddress.FirstAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber,
            this.RangeAddress.LastAddress.RowNumber,
            this.RangeAddress.LastAddress.ColumnNumber
        );

        // Shift formulas first
        foreach (
            XLCell cell in this
                .Worksheet.Workbook.Worksheets.Cast<XLWorksheet>()
                .SelectMany(ws => ws.Internals.CellsCollection.GetCells(c => c.HasFormula))
        )
        {
            if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
            {
                cell.ShiftFormulaRows((XLRange)shiftedRangeFormula, numberOfRows * -1);
            }
            else
            {
                cell.ShiftFormulaColumns((XLRange)shiftedRangeFormula, numberOfColumns * -1);
            }
        }

        // Range to shift...
        int columnModifier = 0;
        int rowModifier = 0;
        Area range = Area.FromRangeAddress(this.RangeAddress);
        switch (shiftDeleteCells)
        {
            case XLShiftDeletedCells.ShiftCellsLeft:
                this.Worksheet.Internals.CellsCollection.DeleteAreaAndShiftLeft(range);
                columnModifier = this.ColumnCount();
                break;

            case XLShiftDeletedCells.ShiftCellsUp:
                this.Worksheet.Internals.CellsCollection.DeleteAreaAndShiftUp(range);
                rowModifier = this.RowCount();
                break;
        }

        List<XLRange> mergesToRemove =
        [
            .. this.Worksheet.Internals.MergedRanges.Where<XLRange>(this.Contains),
        ];
        mergesToRemove.ForEach(r => this.Worksheet.Internals.MergedRanges.Remove(r));

        XLRange shiftedRange = this.AsRange();
        if (shiftDeleteCells == XLShiftDeletedCells.ShiftCellsUp)
        {
            this.Worksheet.NotifyRangeShiftedRows(shiftedRange, rowModifier * -1);
        }
        else
        {
            this.Worksheet.NotifyRangeShiftedColumns(shiftedRange, columnModifier * -1);
        }

        this.Worksheet.DeleteRange(this.RangeAddress);
    }

    public override string ToString() =>
        string.Concat(
            this.Worksheet.Name.EscapeSheetName(),
            '!',
            this.RangeAddress.FirstAddress,
            ':',
            this.RangeAddress.LastAddress
        );

    protected IXLRangeAddress ShiftColumns(
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int columnsShifted
    )
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid)
        {
            return thisRangeAddress;
        }

        bool allRowsAreCovered =
            thisRangeAddress.FirstAddress.RowNumber
                >= shiftedRange.RangeAddress.FirstAddress.RowNumber
            && thisRangeAddress.LastAddress.RowNumber
                <= shiftedRange.RangeAddress.LastAddress.RowNumber;

        if (!allRowsAreCovered)
        {
            return thisRangeAddress;
        }

        bool shiftLeftBoundary =
            (
                columnsShifted > 0
                && thisRangeAddress.FirstAddress.ColumnNumber
                    >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
            )
            || (
                columnsShifted < 0
                && thisRangeAddress.FirstAddress.ColumnNumber
                    > shiftedRange.RangeAddress.FirstAddress.ColumnNumber
            );

        bool shiftRightBoundary =
            thisRangeAddress.LastAddress.ColumnNumber
            >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber;

        int newLeftBoundary = thisRangeAddress.FirstAddress.ColumnNumber;
        if (shiftLeftBoundary)
        {
            if (
                newLeftBoundary + columnsShifted
                > shiftedRange.RangeAddress.FirstAddress.ColumnNumber
            )
            {
                newLeftBoundary = newLeftBoundary + columnsShifted;
            }
            else
            {
                newLeftBoundary = shiftedRange.RangeAddress.FirstAddress.ColumnNumber;
            }
        }

        int newRightBoundary = thisRangeAddress.LastAddress.ColumnNumber;
        if (shiftRightBoundary)
        {
            newRightBoundary += columnsShifted;
        }

        bool destroyedByShift = newRightBoundary < newLeftBoundary;

        XLAddress firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
        XLAddress lastAddress = (XLAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = this.Worksheet.InvalidAddress;
            lastAddress = this.Worksheet.InvalidAddress;
            this.Worksheet.DeleteRange(this.RangeAddress);
        }

        if (shiftLeftBoundary)
        {
            firstAddress = new XLAddress(
                this.Worksheet,
                thisRangeAddress.FirstAddress.RowNumber,
                newLeftBoundary,
                thisRangeAddress.FirstAddress.FixedRow,
                thisRangeAddress.FirstAddress.FixedColumn
            );
        }

        if (shiftRightBoundary)
        {
            lastAddress = new XLAddress(
                this.Worksheet,
                thisRangeAddress.LastAddress.RowNumber,
                newRightBoundary,
                thisRangeAddress.LastAddress.FixedRow,
                thisRangeAddress.LastAddress.FixedColumn
            );
        }

        return new XLRangeAddress(firstAddress, lastAddress);
    }

    protected IXLRangeAddress ShiftRows(
        IXLRangeAddress thisRangeAddress,
        XLRange shiftedRange,
        int rowsShifted
    )
    {
        if (!thisRangeAddress.IsValid || !shiftedRange.RangeAddress.IsValid)
        {
            return thisRangeAddress;
        }

        bool allColumnsAreCovered =
            thisRangeAddress.FirstAddress.ColumnNumber
                >= shiftedRange.RangeAddress.FirstAddress.ColumnNumber
            && thisRangeAddress.LastAddress.ColumnNumber
                <= shiftedRange.RangeAddress.LastAddress.ColumnNumber;

        if (!allColumnsAreCovered)
        {
            return thisRangeAddress;
        }

        bool shiftTopBoundary =
            (
                rowsShifted > 0
                && thisRangeAddress.FirstAddress.RowNumber
                    >= shiftedRange.RangeAddress.FirstAddress.RowNumber
            )
            || (
                rowsShifted < 0
                && thisRangeAddress.FirstAddress.RowNumber
                    > shiftedRange.RangeAddress.FirstAddress.RowNumber
            );

        bool shiftBottomBoundary =
            thisRangeAddress.LastAddress.RowNumber
            >= shiftedRange.RangeAddress.FirstAddress.RowNumber;

        int newTopBoundary = thisRangeAddress.FirstAddress.RowNumber;
        if (shiftTopBoundary)
        {
            if (newTopBoundary + rowsShifted > shiftedRange.RangeAddress.FirstAddress.RowNumber)
            {
                newTopBoundary = newTopBoundary + rowsShifted;
            }
            else
            {
                newTopBoundary = shiftedRange.RangeAddress.FirstAddress.RowNumber;
            }
        }

        int newBottomBoundary = thisRangeAddress.LastAddress.RowNumber;
        if (shiftBottomBoundary)
        {
            newBottomBoundary += rowsShifted;
        }

        bool destroyedByShift = newBottomBoundary < newTopBoundary;

        XLAddress firstAddress = (XLAddress)thisRangeAddress.FirstAddress;
        XLAddress lastAddress = (XLAddress)thisRangeAddress.LastAddress;

        if (destroyedByShift)
        {
            firstAddress = this.Worksheet.InvalidAddress;
            lastAddress = this.Worksheet.InvalidAddress;
            this.Worksheet.DeleteRange(this.RangeAddress);
        }

        if (shiftTopBoundary)
        {
            firstAddress = new XLAddress(
                this.Worksheet,
                newTopBoundary,
                thisRangeAddress.FirstAddress.ColumnNumber,
                thisRangeAddress.FirstAddress.FixedRow,
                thisRangeAddress.FirstAddress.FixedColumn
            );
        }

        if (shiftBottomBoundary)
        {
            lastAddress = new XLAddress(
                this.Worksheet,
                newBottomBoundary,
                thisRangeAddress.LastAddress.ColumnNumber,
                thisRangeAddress.LastAddress.FixedRow,
                thisRangeAddress.LastAddress.FixedColumn
            );
        }

        return new XLRangeAddress(firstAddress, lastAddress);
    }

    public IXLRange RangeUsed() => this.RangeUsed(XLCellsUsedOptions.AllContents);

    public IXLRange RangeUsed(XLCellsUsedOptions options)
    {
        IXLCell firstCell = (this as IXLRangeBase).FirstCellUsed(options);
        if (firstCell == null)
        {
            return null;
        }

        IXLCell lastCell = (this as IXLRangeBase).LastCellUsed(options);
        return this.Worksheet.Range(firstCell, lastCell);
    }

    public virtual void CopyTo(IXLRangeBase target) => this.CopyTo((XLCell)target.FirstCell());

    internal void CopyTo(XLCell target) => target.CopyFrom(this);

    //public IXLChart CreateChart(Int32 firstRow, Int32 firstColumn, Int32 lastRow, Int32 lastColumn)
    //{
    //    IXLChart chart = new XLChartWorksheet;
    //    chart.FirstRow = firstRow;
    //    chart.LastRow = lastRow;
    //    chart.LastColumn = lastColumn;
    //    chart.FirstColumn = firstColumn;
    //    Worksheet.Charts.Add(chart);
    //    return chart;
    //}

    IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, string name) =>
        this.CreatePivotTable(targetCell, name);

    public XLPivotTable CreatePivotTable(IXLCell targetCell, string name) =>
        (XLPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, this.AsRange());

    public virtual IXLAutoFilter SetAutoFilter() => this.SetAutoFilter(true);

    public IXLAutoFilter SetAutoFilter(bool value)
    {
        if (value)
        {
            return this.Worksheet.AutoFilter.Set(this);
        }
        else
        {
            return this.Worksheet.AutoFilter.Clear();
        }
    }

    #region Sort

    public IXLSortElements SortRows => this._sortRows ?? (this._sortRows = new XLSortElements());

    public IXLSortElements SortColumns =>
        this._sortColumns ?? (this._sortColumns = new XLSortElements());

    private string DefaultSortString()
    {
        StringBuilder sb = new();
        int maxColumn = this.ColumnCount();
        if (maxColumn == XlsxSharp.XLHelper.MaxColumnNumber)
        {
            maxColumn = (this as IXLRangeBase)
                .LastCellUsed(XLCellsUsedOptions.All)
                .Address.ColumnNumber;
        }

        for (int i = 1; i <= maxColumn; i++)
        {
            if (sb.Length > 0)
            {
                sb.Append(',');
            }

            sb.Append(i);
        }

        return sb.ToString();
    }

    public IXLRangeBase Sort()
    {
        if (!this.SortColumns.Any())
        {
            return this.Sort(this.DefaultSortString());
        }

        this.SortRangeRows();
        return this;
    }

    public IXLRangeBase Sort(
        string columnsToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        this.SortColumns.Clear();
        if (string.IsNullOrWhiteSpace(columnsToSortBy))
        {
            columnsToSortBy = this.DefaultSortString();
        }

        this.SortColumns.CastTo<XLSortElements>()
            .AddRange(ParseSortOrder(columnsToSortBy, sortOrder, matchCase, ignoreBlanks));

        this.SortRangeRows();
        return this;
    }

    public IXLRangeBase Sort(
        int columnToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    ) => this.Sort(columnToSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);

    public IXLRangeBase SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        this.SortRows.Clear();
        int maxColumn = this.ColumnCount();
        if (maxColumn == XlsxSharp.XLHelper.MaxColumnNumber)
        {
            maxColumn = (this as IXLRangeBase)
                .LastCellUsed(XLCellsUsedOptions.All)
                .Address.ColumnNumber;
        }

        for (int i = 1; i <= maxColumn; i++)
        {
            this.SortRows.Add(i, sortOrder, ignoreBlanks, matchCase);
        }

        this.SortRangeColumns();
        return this;
    }

    private void SortRangeRows()
    {
        Area sortRange = this.SheetRange;
        XLCellsCollection cellsCollection = this.Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireColumn())
        {
            // If we're dealing with the entire column, we're not interested in the unused cells
            int lastRowUsed = cellsCollection.LastRowUsed(Area.Full, XLCellsUsedOptions.Contents);
            sortRange = new Area(
                sortRange.FirstPoint,
                new Point(lastRowUsed, sortRange.RightColumn)
            );
        }

        XLRangeRowsSortComparer comparer = new(this.Worksheet, sortRange, this.SortColumns);
        int[] rows = new int[sortRange.Height];
        for (int i = 0; i < sortRange.Height; ++i)
        {
            rows[i] = i + sortRange.TopRow;
        }

        Array.Sort(rows, comparer);

        cellsCollection.RemapRows(rows, sortRange);
    }

    private void SortRangeColumns()
    {
        Area sortRange = this.SheetRange;
        XLCellsCollection cellsCollection = this.Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireRow())
        {
            // If we're dealing with the entire row, we're not interested in the unused cells
            int lastColumnCell = cellsCollection.LastColumnUsed(
                Area.Full,
                XLCellsUsedOptions.Contents
            );
            sortRange = new Area(
                sortRange.FirstPoint,
                new Point(sortRange.BottomRow, lastColumnCell)
            );
        }

        XLRangeColumnsSortComparer comparer = new(this.Worksheet, sortRange, this.SortRows);
        int[] columns = new int[sortRange.Width];
        for (int i = 0; i < sortRange.Width; ++i)
        {
            columns[i] = i + sortRange.LeftColumn;
        }

        Array.Sort(columns, comparer);

        cellsCollection.RemapColumns(columns, sortRange);
    }

    private static IEnumerable<XLSortElement> ParseSortOrder(
        string columnsToSortBy,
        XLSortOrder defaultSortOrder,
        bool matchCase,
        bool ignoreBlanks
    )
    {
        foreach (string sortColumn in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            XLSortOrder sortOrder = defaultSortOrder;

            string columnName;
            if (sortColumn.Contains(' '))
            {
                string[] pair = sortColumn.Split(' ');
                columnName = pair[0];
                sortOrder = pair[1].Equals("ASC", StringComparison.OrdinalIgnoreCase)
                    ? XLSortOrder.Ascending
                    : XLSortOrder.Descending;
            }
            else
            {
                columnName = sortColumn;
            }

            if (!int.TryParse(columnName, out int columnNumber))
            {
                columnNumber = XlsxSharp.XLHelper.GetColumnNumberFromLetter(columnName);
            }

            yield return new XLSortElement(columnNumber, sortOrder, ignoreBlanks, matchCase);
        }
    }

    #endregion Sort

    public XLRangeColumn ColumnQuick(int column)
    {
        XLAddress firstCellAddress = new(
            this.Worksheet,
            this.RangeAddress.FirstAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber + column - 1,
            false,
            false
        );
        XLAddress lastCellAddress = new(
            this.Worksheet,
            this.RangeAddress.LastAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber + column - 1,
            false,
            false
        );
        return this.Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
    }

    public IXLConditionalFormat AddConditionalFormat()
    {
        XLConditionalFormat cf = new(this.Worksheet, this.SheetRange.ToAreaList());
        this.Worksheet.ConditionalFormats.Add(cf);
        return cf;
    }

    public void Select() => this.Worksheet.SelectedRanges.Add(this.AsRange());

    public IXLRangeBase Grow() => this.Grow(1);

    public IXLRangeBase Grow(int growCount)
    {
        int firstRow = Math.Max(1, this.RangeAddress.FirstAddress.RowNumber - growCount);
        int firstColumn = Math.Max(1, this.RangeAddress.FirstAddress.ColumnNumber - growCount);

        int lastRow = Math.Min(
            XlsxSharp.XLHelper.MaxRowNumber,
            this.RangeAddress.LastAddress.RowNumber + growCount
        );
        int lastColumn = Math.Min(
            XlsxSharp.XLHelper.MaxColumnNumber,
            this.RangeAddress.LastAddress.ColumnNumber + growCount
        );

        return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    public IXLRangeBase Shrink() => this.Shrink(1);

    public IXLRangeBase Shrink(int shrinkCount)
    {
        int firstRow = this.RangeAddress.FirstAddress.RowNumber + shrinkCount;
        int firstColumn = this.RangeAddress.FirstAddress.ColumnNumber + shrinkCount;

        int lastRow = this.RangeAddress.LastAddress.RowNumber - shrinkCount;
        int lastColumn = this.RangeAddress.LastAddress.ColumnNumber - shrinkCount;

        if (firstRow > lastRow || firstColumn > lastColumn)
        {
            return null;
        }

        return this.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    public IXLRangeAddress Intersection(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    )
    {
        if (otherRange == null)
        {
            return null;
        }

        if (!this.Worksheet.Equals(otherRange.Worksheet))
        {
            return null;
        }

        if (thisRangePredicate == null && otherRangePredicate == null)
        {
            // Special case, no predicates. We can optimise this a bit then.
            return this.RangeAddress.Intersection(otherRange.RangeAddress);
        }
        else
        {
            thisRangePredicate = thisRangePredicate ?? (c => true);
            otherRangePredicate = otherRangePredicate ?? (c => true);

            IXLCells intersectionCells = this.Cells(c =>
                thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c)
            );

            if (!intersectionCells.Any())
            {
                return null;
            }

            int firstRow = intersectionCells.Min(c => c.Address.RowNumber);
            int firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

            int lastRow = intersectionCells.Max(c => c.Address.RowNumber);
            int lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

            return new XLRangeAddress(
                new XLAddress(
                    this.Worksheet,
                    firstRow,
                    firstColumn,
                    fixedRow: false,
                    fixedColumn: false
                ),
                new XLAddress(
                    this.Worksheet,
                    lastRow,
                    lastColumn,
                    fixedRow: false,
                    fixedColumn: false
                )
            );
        }
    }

    public IXLCells SurroundingCells(Func<IXLCell, bool> predicate = null)
    {
        XLCells cells = new(this.Worksheet, false, XLCellsUsedOptions.AllContents, predicate);
        this.Grow().Cells(c => !this.Contains(c)).ForEach(c => cells.Add(c as XLCell));
        return cells;
    }

    public IXLCells Union(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    )
    {
        if (otherRange == null)
        {
            return this.Cells(thisRangePredicate);
        }

        XLCells cells = new(this.Worksheet, false, XLCellsUsedOptions.AllContents);
        if (!this.Worksheet.Equals(otherRange.Worksheet))
        {
            return cells;
        }

        if (thisRangePredicate == null)
        {
            thisRangePredicate = c => true;
        }

        if (otherRangePredicate == null)
        {
            otherRangePredicate = c => true;
        }

        this.Cells(thisRangePredicate)
            .Concat(otherRange.Cells(otherRangePredicate))
            .Distinct()
            .ForEach(c => cells.Add(c as XLCell));
        return cells;
    }

    public IXLCells Difference(
        IXLRangeBase otherRange,
        Func<IXLCell, bool> thisRangePredicate = null,
        Func<IXLCell, bool> otherRangePredicate = null
    )
    {
        if (otherRange == null)
        {
            return this.Cells(thisRangePredicate);
        }

        XLCells cells = new(this.Worksheet, false, XLCellsUsedOptions.AllContents);
        if (!this.Worksheet.Equals(otherRange.Worksheet))
        {
            return cells;
        }

        if (thisRangePredicate == null)
        {
            thisRangePredicate = c => true;
        }

        if (otherRangePredicate == null)
        {
            otherRangePredicate = c => true;
        }

        this.Cells(c => thisRangePredicate(c) && !otherRange.Cells(otherRangePredicate).Contains(c))
            .ForEach(c => cells.Add(c as XLCell));
        return cells;
    }

    private IEnumerable<IXLCell> CellsUsedInternal(
        XLCellsUsedOptions options,
        Func<IXLRange, IXLCell> selector,
        Func<IXLCell, bool> predicate
    )
    {
        predicate ??= (t => true);

        // To avoid unnecessary initialization of thousands cells to not hang on very large CFs, DVs or merged ranges.
        XLCellsUsedOptions opt =
            options
            & ~XLCellsUsedOptions.ConditionalFormats
            & ~XLCellsUsedOptions.DataValidation
            & ~XLCellsUsedOptions.MergedRanges;

        // If opt == 0 then we're basically back at unconstrained, so just set back the original options
        if (opt == XLCellsUsedOptions.NoConstraints)
        {
            opt = options;
        }

        IEnumerable<IXLCell> cellsUsed = this.CellsUsed(opt, predicate);

        if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
        {
            Area area = this.SheetRange;
            cellsUsed = cellsUsed.Union(
                this.Worksheet.ConditionalFormats.SelectMany<XLConditionalFormat, Area>(cf =>
                        cf.Areas.IntersectingWith(area)
                    )
                    .Select(cfArea => this.Worksheet.Range(cfArea))
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XLCellsUsedOptions.DataValidation))
        {
            cellsUsed = cellsUsed.Union(
                this.Worksheet.DataValidations.GetAllInRange(this.RangeAddress)
                    .SelectMany(dv => dv.Ranges)
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XLCellsUsedOptions.MergedRanges))
        {
            cellsUsed = cellsUsed.Union(
                this.Worksheet.MergedRanges.GetIntersectedRanges(this.RangeAddress)
                    .Select(selector)
                    .Where(predicate)
            );
        }

        return cellsUsed;
    }
}

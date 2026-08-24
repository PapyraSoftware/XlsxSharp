using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Extensions;
using XlsxSharp.Graphics;

namespace XlsxSharp.Excel.Rows;

internal sealed class XLRow : XLRangeBase, IXLRow, IXLFormatContainer
{
    /// <summary>
    /// Don't use directly, use properties.
    /// </summary>
    private XlRowFlags _flags;
    private double _height;
    private int _outlineLevel;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLRow(XLWorksheet worksheet, int row)
        : base(XLRangeAddress.EntireRow(worksheet, row))
    {
        this.SetRowNumber(row);

        this._height = worksheet.RowHeight;
    }

    internal XLRowArea Area => new(this.Worksheet.Name, this.RowNumber());

    public override XLRangeType RangeType => XLRangeType.Row;

    public bool Collapsed
    {
        get => this._flags.HasFlag(XlRowFlags.Collapsed);
        set
        {
            if (value)
            {
                this._flags |= XlRowFlags.Collapsed;
            }
            else
            {
                this._flags &= ~XlRowFlags.Collapsed;
            }
        }
    }

    /// <summary>
    /// Distance in pixels from the bottom of the cells in the current row to the typographical
    /// baseline of the cell content if, hypothetically, the zoom level for the sheet containing
    /// this row is 100 percent and the cell has bottom-alignment formatting.
    /// </summary>
    /// <remarks>
    /// If the attribute is set, it sets customHeight to true even if the customHeight is explicitly
    /// set to false. Custom height means no auto-sizing by Excel on load, so if row has this
    /// attribute, it stops Excel from auto-sizing the height of a row to fit the content on load.
    /// </remarks>
    public double? DyDescent { get; set; }

    /// <summary>
    /// Should cells in the row display phonetic? This doesn't actually affect whether the phonetic are
    /// shown in the row, that depends entirely on the <see cref="IXLCell.ShowPhonetic"/> property
    /// of a cell. This property determines whether a new cell in the row will have it's phonetic turned on
    /// (and also the state of the "Show or hide phonetic" in Excel when whole row is selected).
    /// Default is <c>false</c>.
    /// </summary>
    public bool ShowPhonetic
    {
        get => this._flags.HasFlag(XlRowFlags.ShowPhonetic);
        set
        {
            if (value)
            {
                this._flags |= XlRowFlags.ShowPhonetic;
            }
            else
            {
                this._flags &= ~XlRowFlags.ShowPhonetic;
            }
        }
    }

    public bool Loading
    {
        get => this._flags.HasFlag(XlRowFlags.Loading);
        set
        {
            if (value)
            {
                this._flags |= XlRowFlags.Loading;
            }
            else
            {
                this._flags &= ~XlRowFlags.Loading;
            }
        }
    }

    /// <summary>
    /// Does row have an individual height or is it derived from the worksheet <see cref="XLWorksheet.RowHeight"/>?
    /// </summary>
    public bool HeightChanged
    {
        get => this._flags.HasFlag(XlRowFlags.HeightChanged);
        private set
        {
            if (value)
            {
                this._flags |= XlRowFlags.HeightChanged;
            }
            else
            {
                this._flags &= ~XlRowFlags.HeightChanged;
            }
        }
    }

    #region IXLRow Members

    public double Height
    {
        get => this._height;
        set
        {
            if (!this.Loading)
            {
                this.HeightChanged = true;
            }

            this._height = value;
        }
    }

    IXLCells IXLRow.Cells(string cellsInRow) => this.Cells(cellsInRow);

    IXLCells IXLRow.Cells(int firstColumn, int lastColumn) => this.Cells(firstColumn, lastColumn);

    public void ClearHeight()
    {
        this.Height = this.Worksheet.RowHeight;
        this.HeightChanged = false;
    }

    public void Delete()
    {
        int rowNumber = this.RowNumber();
        this.AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
        this.Worksheet.DeleteRow(rowNumber);
    }

    public new IXLRows InsertRowsBelow(int numberOfRows)
    {
        int rowNum = this.RowNumber();
        this.Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
        XLRange asRange = this.Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsBelowVoid(true, numberOfRows);

        IXLRows newRows = this.Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);
        foreach (IXLRow newRow in newRows)
        {
            XLRow internalRow = this.Worksheet.Internals.RowsCollection[newRow.RowNumber()];
            internalRow._height = this.Height;
            internalRow.FormatValue = this.FormatValue; // Is within a worbook
            internalRow.Collapsed = this.Collapsed;
            internalRow.IsHidden = this.IsHidden;
            internalRow._outlineLevel = this.OutlineLevel;
        }

        return newRows;
    }

    public new IXLRows InsertRowsAbove(int numberOfRows)
    {
        int rowNum = this.RowNumber();
        if (rowNum > 1)
        {
            return this.Worksheet.Row(rowNum - 1).InsertRowsBelow(numberOfRows);
        }

        this.Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
        XLRange asRange = this.Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsAboveVoid(true, numberOfRows);

        return this.Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
    }

    public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLCell Cell(int columnNumber) => this.Cell(1, columnNumber);

    public override XLCell Cell(string columnLetter) => this.Cell(1, columnLetter);

    IXLCell IXLRow.Cell(string columnLetter) => this.Cell(columnLetter);

    public override IXLCells Cells() => this.Cells(true, XLCellsUsedOptions.All);

    public override XLCells Cells(bool usedCellsOnly)
    {
        if (usedCellsOnly)
        {
            return this.Cells(true, XLCellsUsedOptions.AllContents);
        }
        else
        {
            return this.Cells(
                this.FirstCellUsed().Address.ColumnNumber,
                this.LastCellUsed().Address.ColumnNumber
            );
        }
    }

    public override XLCells Cells(string cellsInRow)
    {
        XLCells retVal = new(this.Worksheet, false, XLCellsUsedOptions.AllContents);
        string[] rangePairs = cellsInRow.Split(',');
        foreach (string pair in rangePairs)
        {
            retVal.Add(this.Range(pair.Trim()).RangeAddress);
        }

        return retVal;
    }

    public XLCells Cells(int firstColumn, int lastColumn) =>
        this.Cells(firstColumn + ":" + lastColumn);

    public IXLCells Cells(string firstColumn, string lastColumn) =>
        this.Cells(
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(firstColumn)
                + ":"
                + XlsxSharp.XLHelper.GetColumnNumberFromLetter(lastColumn)
        );

    public IXLRow AdjustToContents(int startColumn) =>
        this.AdjustToContents(startColumn, XlsxSharp.XLHelper.MaxColumnNumber);

    public IXLRow AdjustToContents(int startColumn, int endColumn) =>
        this.AdjustToContents(startColumn, endColumn, 0, double.MaxValue);

    public IXLRow AdjustToContents(double minHeight, double maxHeight) =>
        this.AdjustToContents(1, XlsxSharp.XLHelper.MaxColumnNumber, minHeight, maxHeight);

    public IXLRow AdjustToContents(int startColumn, double minHeight, double maxHeight) =>
        this.AdjustToContents(
            startColumn,
            XlsxSharp.XLHelper.MaxColumnNumber,
            minHeight,
            maxHeight
        );

    public IXLRow AdjustToContents(
        int startColumn,
        int endColumn,
        double minHeightPt,
        double maxHeightPt
    )
    {
        IXLGraphicEngine engine = this.Worksheet.Workbook.GraphicEngine;
        Dpi dpi = new(this.Worksheet.Workbook.DpiX, this.Worksheet.Workbook.DpiY);

        int rowHeightPx = this.CalculateMinRowHeight(startColumn, endColumn, engine, dpi);

        double rowHeightPt = XlsxSharp.XLHelper.PixelsToPoints(rowHeightPx, dpi.Y);
        if (rowHeightPt <= 0)
        {
            rowHeightPt = this.Worksheet.RowHeight;
        }

        if (minHeightPt > rowHeightPt)
        {
            rowHeightPt = minHeightPt;
        }

        if (maxHeightPt < rowHeightPt)
        {
            rowHeightPt = maxHeightPt;
        }

        this.Height = rowHeightPt;

        return this;
    }

    private int CalculateMinRowHeight(
        int startColumn,
        int endColumn,
        IXLGraphicEngine engine,
        Dpi dpi
    )
    {
        List<GlyphBox> glyphs = [];
        int rowHeightPx = 0;
        foreach (XLCell cell in this.Row(startColumn, endColumn).CellsUsed().Cast<XLCell>())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
            {
                continue;
            }

            XLCellFormatValue cellFormat = this.Worksheet.GetStyleValue(cell.Point);

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            int cellHeightPx = (int)
                Math.Ceiling(GetContentHeight(cellFormat.Alignment.TextRotation.Value, glyphs));

            rowHeightPx = Math.Max(cellHeightPx, rowHeightPx);
        }

        return rowHeightPx;
    }

    private static double GetContentHeight(int textRotationDeg, List<GlyphBox> glyphs)
    {
        if (textRotationDeg == 0)
        {
            double textHeight = 0d;
            double lineMaxHeight = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    float cellHeightPx = glyph.LineHeight;
                    lineMaxHeight = Math.Max(cellHeightPx, lineMaxHeight);
                }
                else
                {
                    // At the end of each line, add height of the line to total height.
                    textHeight += lineMaxHeight;
                    lineMaxHeight = 0d;
                }
            }

            // If the last line ends without EOL, it must be also counted
            textHeight += lineMaxHeight;

            return textHeight;
        }
        else if (textRotationDeg == 255)
        {
            // Glyphs are vertically aligned.
            float textHeight = glyphs.Sum(static g => g.LineHeight);
            return textHeight;
        }
        else
        {
            // Rotated text
            double width = 0d;
            double height = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                width += glyph.AdvanceWidth;
                height = Math.Max(glyph.LineHeight, height);
            }

            double projectedWidth = Math.Sin(XlsxSharp.XLHelper.DegToRad(textRotationDeg)) * width;
            double projectedHeight =
                Math.Cos(XlsxSharp.XLHelper.DegToRad(textRotationDeg)) * height;
            return projectedWidth + projectedHeight;
        }
    }

    public IXLRow Hide()
    {
        this.IsHidden = true;
        return this;
    }

    public IXLRow Unhide()
    {
        this.IsHidden = false;
        return this;
    }

    public bool IsHidden
    {
        get => this._flags.HasFlag(XlRowFlags.IsHidden);
        set
        {
            if (value)
            {
                this._flags |= XlRowFlags.IsHidden;
            }
            else
            {
                this._flags &= ~XlRowFlags.IsHidden;
            }
        }
    }

    public int OutlineLevel
    {
        get => this._outlineLevel;
        set
        {
            if (value < 0 || value > 8)
            {
                throw new ArgumentOutOfRangeException(
                    "value",
                    "Outline level must be between 0 and 8."
                );
            }

            this.Worksheet.IncrementColumnOutline(value);
            this.Worksheet.DecrementColumnOutline(this._outlineLevel);
            this._outlineLevel = value;
        }
    }

    public IXLRow Group() => this.Group(false);

    public IXLRow Group(int outlineLevel) => this.Group(outlineLevel, false);

    public IXLRow Ungroup() => this.Ungroup(false);

    public IXLRow Group(bool collapse)
    {
        if (this.OutlineLevel < 8)
        {
            this.OutlineLevel += 1;
        }

        this.Collapsed = collapse;
        return this;
    }

    public IXLRow Group(int outlineLevel, bool collapse)
    {
        this.OutlineLevel = outlineLevel;
        this.Collapsed = collapse;
        return this;
    }

    public IXLRow Ungroup(bool ungroupFromAll)
    {
        if (ungroupFromAll)
        {
            this.OutlineLevel = 0;
        }
        else
        {
            if (this.OutlineLevel > 0)
            {
                this.OutlineLevel -= 1;
            }
        }
        return this;
    }

    public IXLRow Collapse()
    {
        this.Collapsed = true;
        return this.Hide();
    }

    public IXLRow Expand()
    {
        this.Collapsed = false;
        return this.Unhide();
    }

    public int CellCount() =>
        this.RangeAddress.LastAddress.ColumnNumber
        - this.RangeAddress.FirstAddress.ColumnNumber
        + 1;

    public new IXLRow Sort() => this.SortLeftToRight();

    public new IXLRow SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXLRangeRow IXLRow.CopyTo(IXLCell target)
    {
        IXLRange copy = this.AsRange().CopyTo(target);
        return copy.Row(1);
    }

    IXLRangeRow IXLRow.CopyTo(IXLRangeBase target)
    {
        IXLRange copy = this.AsRange().CopyTo(target);
        return copy.Row(1);
    }

    public IXLRow CopyTo(IXLRow row)
    {
        row.Clear();
        XLRow newRow = (XLRow)row;
        newRow._height = this._height;
        newRow.HeightChanged = this.HeightChanged;
        newRow.IsHidden = this.IsHidden;
        if (this.FormatValue is not null)
        {
            newRow.FormatValue = newRow.Worksheet.Workbook.Styles.GetRegisteredCellFormat(
                this.GetFormat()
            );
        }

        this.AsRange().CopyTo(row);

        return newRow;
    }

    public IXLRangeRow Row(int start, int end) => this.Range(1, start, 1, end).Row(1);

    public IXLRangeRow Row(IXLCell start, IXLCell end) =>
        this.Row(start.Address.ColumnNumber, end.Address.ColumnNumber);

    public IXLRangeRows Rows(string rows)
    {
        XLRangeRows retVal = new(this.Worksheet);
        string[] rowPairs = rows.Split(',');
        foreach (string pair in rowPairs)
        {
            this.AsRange().Rows(pair.Trim()).ForEach(retVal.Add);
        }

        return retVal;
    }

    public IXLRow AddHorizontalPageBreak()
    {
        this.Worksheet.PageSetup.AddHorizontalPageBreak(this.RowNumber());
        return this;
    }

    public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents) =>
        this.Row(
            (this as IXLRangeBase).FirstCellUsed(options),
            (this as IXLRangeBase).LastCellUsed(options)
        );

    #endregion IXLRow Members


    #region IXLFormatContainer

    /// <remarks>
    /// Format of a row or <c>null</c> for not defined format.
    /// </remarks>
    /// <inheritdoc cref="IXLFormatContainer.FormatValue"/>
    public XLCellFormatValue? FormatValue { get; set; }

    internal override XLCellFormat Format => XLCellFormat.ForRow(this);

    private XLCellFormatValue GetFormat() =>
        this.FormatValue ?? this.Worksheet.Workbook.Styles.DefaultCellFormat;

    #endregion

    public override XLRange AsRange() => this.Range(1, 1, 1, XlsxSharp.XLHelper.MaxColumnNumber);

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        //do nothing
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        // rows are shifted by XLRowCollection
    }

    internal void SetRowNumber(int row) =>
        this.RangeAddress = new XLRangeAddress(
            new XLAddress(
                this.Worksheet,
                row,
                1,
                this.RangeAddress.FirstAddress.FixedRow,
                this.RangeAddress.FirstAddress.FixedColumn
            ),
            new XLAddress(
                this.Worksheet,
                row,
                XlsxSharp.XLHelper.MaxColumnNumber,
                this.RangeAddress.LastAddress.FixedRow,
                this.RangeAddress.LastAddress.FixedColumn
            )
        );

    public override XLRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
            {
                rangeAddressStr = rangeAddressStr.Replace('-', ':');
            }

            string[] arrRange = rangeAddressStr.Split(':');
            string firstPart = arrRange[0];
            string secondPart = arrRange[1];
            rangeAddressToUse =
                this.FixRowAddress(firstPart) + ":" + this.FixRowAddress(secondPart);
        }
        else
        {
            rangeAddressToUse = this.FixRowAddress(rangeAddressStr);
        }

        XLRangeAddress rangeAddress = new(this.Worksheet, rangeAddressToUse);
        return this.Range(rangeAddress);
    }

    public IXLRow AdjustToContents() => this.AdjustToContents(1);

    private XLRow RowShift(int rowsToShift) => this.Worksheet.Row(this.RowNumber() + rowsToShift);

    #region XLRow Above

    IXLRow IXLRow.RowAbove() => this.RowAbove();

    IXLRow IXLRow.RowAbove(int step) => this.RowAbove(step);

    public XLRow RowAbove() => this.RowAbove(1);

    public XLRow RowAbove(int step) => this.RowShift(step * -1);

    #endregion XLRow Above

    #region XLRow Below

    IXLRow IXLRow.RowBelow() => this.RowBelow();

    IXLRow IXLRow.RowBelow(int step) => this.RowBelow(step);

    public XLRow RowBelow() => this.RowBelow(1);

    public XLRow RowBelow(int step) => this.RowShift(step);

    #endregion XLRow Below

    public override bool IsEmpty() => this.IsEmpty(XLCellsUsedOptions.AllContents);

    public override bool IsEmpty(XLCellsUsedOptions options)
    {
        if (options.HasFlag(XLCellsUsedOptions.NormalFormats) && this.FormatValue is not null)
        {
            return false;
        }

        return base.IsEmpty(options);
    }

    public override bool IsEntireRow() => true;

    public override bool IsEntireColumn() => false;

    /// <summary>
    /// Flag enum to save space, instead of wasting byte for each flag.
    /// </summary>
    [Flags]
    private enum XlRowFlags : byte
    {
        Collapsed = 1,
        IsHidden = 2,
        ShowPhonetic = 4,
        HeightChanged = 8,
        Loading = 16,
    }
}

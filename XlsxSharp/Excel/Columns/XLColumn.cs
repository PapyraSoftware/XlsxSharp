using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.Graphics;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel;

internal class XLColumn : XLRangeBase, IXLColumn, IXLFormatContainer
{
    private int _outlineLevel;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLColumn(XLWorksheet worksheet, int column)
        : base(XLRangeAddress.EntireColumn(worksheet, column))
    {
        this.SetColumnNumber(column);

        this.Width = worksheet.ColumnWidth;
    }

    /// <summary>
    /// Get area of this column.
    /// </summary>
    internal XLColumnArea Area => new(this.Worksheet.Name, this.ColumnNumber());

    public override XLRangeType RangeType => XLRangeType.Column;

    public bool Collapsed { get; set; }

    #region IXLColumn Members

    public double Width { get; set; }

    IXLCells IXLColumn.Cells(string cellsInColumn) => this.Cells(cellsInColumn);

    IXLCells IXLColumn.Cells(int firstRow, int lastRow) => this.Cells(firstRow, lastRow);

    public void Delete()
    {
        int columnNumber = this.ColumnNumber();
        this.Delete(XLShiftDeletedCells.ShiftCellsLeft);
        this.Worksheet.DeleteColumn(columnNumber);
    }

    public new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLCell Cell(int rowNumber) => this.Cell(rowNumber, 1);

    public override XLCells Cells(string cellsInColumn)
    {
        XLCells retVal = new(this.Worksheet, false, XLCellsUsedOptions.All);
        string[] rangePairs = cellsInColumn.Split(',');
        foreach (string pair in rangePairs)
        {
            retVal.Add(this.Range(pair.Trim()).RangeAddress);
        }

        return retVal;
    }

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
                this.FirstCellUsed().Address.RowNumber,
                this.LastCellUsed().Address.RowNumber
            );
        }
    }

    public XLCells Cells(int firstRow, int lastRow) => this.Cells(firstRow + ":" + lastRow);

    public new IXLColumns InsertColumnsAfter(int numberOfColumns)
    {
        int columnNum = this.ColumnNumber();
        this.Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(
            columnNum + 1,
            numberOfColumns
        );
        this.Worksheet.Column(columnNum).InsertColumnsAfterVoid(true, numberOfColumns);
        IXLColumns newColumns = this.Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
        foreach (IXLColumn newColumn in newColumns)
        {
            XLColumn internalColumn = this.Worksheet.Internals.ColumnsCollection[
                newColumn.ColumnNumber()
            ];
            internalColumn.Width = this.Width;
            internalColumn.FormatValue = this.FormatValue; // Is within a worbook
            internalColumn.Collapsed = this.Collapsed;
            internalColumn.IsHidden = this.IsHidden;
            internalColumn._outlineLevel = this.OutlineLevel;
        }

        return newColumns;
    }

    public new IXLColumns InsertColumnsBefore(int numberOfColumns)
    {
        int columnNum = this.ColumnNumber();
        if (columnNum > 1)
        {
            return this.Worksheet.Column(columnNum - 1).InsertColumnsAfter(numberOfColumns);
        }

        this.Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
        this.Worksheet.Column(columnNum).InsertColumnsBeforeVoid(true, numberOfColumns);

        return this.Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
    }

    public IXLColumn AdjustToContents() => this.AdjustToContents(1);

    public IXLColumn AdjustToContents(int startRow) =>
        this.AdjustToContents(startRow, XlsxSharp.XLHelper.MaxRowNumber);

    public IXLColumn AdjustToContents(int startRow, int endRow) =>
        this.AdjustToContents(startRow, endRow, 0, double.MaxValue);

    public IXLColumn AdjustToContents(double minWidth, double maxWidth) =>
        this.AdjustToContents(1, XlsxSharp.XLHelper.MaxRowNumber, minWidth, maxWidth);

    public IXLColumn AdjustToContents(int startRow, double minWidth, double maxWidth) =>
        this.AdjustToContents(startRow, XlsxSharp.XLHelper.MaxRowNumber, minWidth, maxWidth);

    public IXLColumn AdjustToContents(
        int startRow,
        int endRow,
        double minWidthNoC,
        double maxWidthNoC
    )
    {
        IXLGraphicEngine engine = this.Worksheet.Workbook.GraphicEngine;
        Dpi dpi = new(this.Worksheet.Workbook.DpiX, this.Worksheet.Workbook.DpiY);
        int columnWidthPx = this.CalculateMinColumnWidth(startRow, endRow, engine, dpi);

        // Maximum digit width, rounded to pixels, so Calibri at 11 pts returns 7 pixels MDW (the correct value)
        int mdw = (int)
            Math.Round(engine.GetMaxDigitWidth(this.Worksheet.Workbook.Format.Font, dpi.X));

        double minWidthInPx = Math.Ceiling(XlsxSharp.XLHelper.NoCToPixels(minWidthNoC, mdw));
        if (columnWidthPx < minWidthInPx)
        {
            columnWidthPx = (int)minWidthInPx;
        }

        double maxWidthInPx = Math.Ceiling(XlsxSharp.XLHelper.NoCToPixels(maxWidthNoC, mdw));
        if (columnWidthPx > maxWidthInPx)
        {
            columnWidthPx = (int)maxWidthInPx;
        }

        double colMaxWidth = XlsxSharp.XLHelper.PixelToNoC(columnWidthPx, mdw);

        // If there is nothing in the column, use worksheet column width.
        if (colMaxWidth <= 0)
        {
            colMaxWidth = this.Worksheet.ColumnWidth;
        }

        this.Width = colMaxWidth;

        return this;
    }

    /// <summary>
    /// Calculate column width in pixels according to the content of cells.
    /// </summary>
    /// <param name="startRow">First row number whose content is used for determination.</param>
    /// <param name="endRow">Last row number whose content is used for determination.</param>
    /// <param name="engine">Engine to determine size of glyphs.</param>
    /// <param name="dpi">DPI of the worksheet.</param>
    private int CalculateMinColumnWidth(int startRow, int endRow, IXLGraphicEngine engine, Dpi dpi)
    {
        List<int> autoFilterRows = [];
        if (this.Worksheet.AutoFilter != null && this.Worksheet.AutoFilter.Range != null)
        {
            autoFilterRows.Add(this.Worksheet.AutoFilter.Range.FirstRow().RowNumber());
        }

        autoFilterRows.AddRange(
            this.Worksheet.Tables.Where<XLTable>(t =>
                    t.AutoFilter != null
                    && t.AutoFilter.Range != null
                    && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow().RowNumber())
                )
                .Select(t => t.AutoFilter.Range.FirstRow().RowNumber())
        );

        // Cache MDW for each font to avoid too many allocations
        Dictionary<XLFontFormatValue, double> scaledMdwMap = new(
            ReferenceEqualityComparer<XLFontFormatValue>.Instance
        );

        // Reusable buffer
        List<GlyphBox> glyphs = [];
        int columnWidthPx = 0;
        foreach (XLCell cell in this.Column(startRow, endRow).CellsUsed())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
            {
                continue;
            }

            XLCellFormatValue cellStyle = this.Worksheet.GetStyleValue(cell.Point);

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            int textWidthPx = (int)
                Math.Ceiling(GetContentWidth(cellStyle.Alignment.TextRotation.Value, glyphs));

            if (!scaledMdwMap.TryGetValue(cellStyle.Font, out double scaledMdw))
            {
                double mdw = engine.GetMaxDigitWidth(cellStyle.Font.ToFontBase(), dpi.X);
                scaledMdw = Math.Round(mdw, MidpointRounding.AwayFromZero);
                scaledMdwMap.Add(cellStyle.Font, scaledMdw);
            }

            // Not sure about rounding, but larger is probably better, so use ceiling.
            // Due to mismatched rendering, add 3% instead of 1.75%, to have additional space.
            int oneSidePadding = (int)Math.Ceiling(textWidthPx * 0.03 + scaledMdw / 4);

            // Cell width if calculated as content width + padding on each side of a content.
            // The one side padding is roughly 1.75% of content + MDW/4.
            // The additional pixel is there for lines between cells.
            int cellWidthPx = textWidthPx + 2 * oneSidePadding + 1;

            if (autoFilterRows.Contains(cell.Address.RowNumber))
            {
                // Autofilter arrow is 16px at 96dpi, scaling through DPI, e.g. 20px at 120dpi
                cellWidthPx += (int)Math.Round(16d * dpi.X / 96d, MidpointRounding.AwayFromZero);
            }

            columnWidthPx = Math.Max(cellWidthPx, columnWidthPx);
        }

        return columnWidthPx;
    }

    private static double GetContentWidth(int textRotationDeg, List<GlyphBox> glyphs)
    {
        if (textRotationDeg == 0)
        {
            double maxTextWidth = 0d;
            double lineTextWidth = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    lineTextWidth += glyph.AdvanceWidth;
                    maxTextWidth = Math.Max(lineTextWidth, maxTextWidth);
                }
                else
                {
                    lineTextWidth = 0;
                }
            }

            return maxTextWidth;
        }
        if (textRotationDeg == 255)
        {
            // Glyphs are arranged vertically, top to bottom.
            double maxGlyphWidth = 0d;
            foreach (GlyphBox grapheme in glyphs)
            {
                maxGlyphWidth = Math.Max(grapheme.AdvanceWidth, maxGlyphWidth);
            }

            return maxGlyphWidth;
        }
        else
        {
            // Glyphs are rotated
            if (textRotationDeg > 90)
            {
                textRotationDeg = 90 - textRotationDeg;
            }

            double totalWidth = 0d;
            double maxHeight = 0d;
            foreach (GlyphBox glyph in glyphs)
            {
                totalWidth += glyph.AdvanceWidth;
                maxHeight = Math.Max(maxHeight, glyph.LineHeight);
            }

            double projectedHeight =
                maxHeight * Math.Cos(XlsxSharp.XLHelper.DegToRad(90 - textRotationDeg));
            double projectedWidth =
                totalWidth * Math.Cos(XlsxSharp.XLHelper.DegToRad(textRotationDeg));
            return projectedWidth + projectedHeight;
        }
    }

    public IXLColumn Hide()
    {
        this.IsHidden = true;
        return this;
    }

    public IXLColumn Unhide()
    {
        this.IsHidden = false;
        return this;
    }

    public bool IsHidden { get; set; }

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

    public IXLColumn Group() => this.Group(false);

    public IXLColumn Group(bool collapse)
    {
        if (this.OutlineLevel < 8)
        {
            this.OutlineLevel += 1;
        }

        this.Collapsed = collapse;
        return this;
    }

    public IXLColumn Group(int outlineLevel) => this.Group(outlineLevel, false);

    public IXLColumn Group(int outlineLevel, bool collapse)
    {
        this.OutlineLevel = outlineLevel;
        this.Collapsed = collapse;
        return this;
    }

    public IXLColumn Ungroup() => this.Ungroup(false);

    public IXLColumn Ungroup(bool ungroupFromAll)
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

    public IXLColumn Collapse()
    {
        this.Collapsed = true;
        return this.Hide();
    }

    public IXLColumn Expand()
    {
        this.Collapsed = false;
        return this.Unhide();
    }

    public int CellCount() =>
        this.RangeAddress.LastAddress.ColumnNumber
        - this.RangeAddress.FirstAddress.ColumnNumber
        + 1;

    public IXLColumn Sort(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    )
    {
        this.Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXLRangeColumn IXLColumn.Column(int start, int end) => this.Column(start, end);

    IXLRangeColumn IXLColumn.CopyTo(IXLCell target)
    {
        IXLRange copy = this.AsRange().CopyTo(target);
        return copy.Column(1);
    }

    IXLRangeColumn IXLColumn.CopyTo(IXLRangeBase target)
    {
        IXLRange copy = this.AsRange().CopyTo(target);
        return copy.Column(1);
    }

    public IXLColumn CopyTo(IXLColumn column)
    {
        column.Clear();
        XLColumn newColumn = (XLColumn)column;
        newColumn.Width = this.Width;
        newColumn.FormatValue = newColumn.Worksheet.Workbook.Styles.GetRegisteredCellFormat(
            this.GetFormat()
        );
        newColumn.IsHidden = this.IsHidden;

        (this as XLRangeBase).CopyTo(column);

        return newColumn;
    }

    public XLRangeColumn Column(int start, int end) => this.Range(start, 1, end, 1).Column(1);

    public IXLRangeColumn Column(IXLCell start, IXLCell end) =>
        this.Column(start.Address.RowNumber, end.Address.RowNumber);

    public IXLRangeColumns Columns(string columns)
    {
        XLRangeColumns retVal = new(this.Worksheet);
        string[] columnPairs = columns.Split(',');
        foreach (string pair in columnPairs)
        {
            this.AsRange().Columns(pair.Trim()).ForEach(retVal.Add);
        }

        return retVal;
    }

    /// <summary>
    ///   Adds a vertical page break after this column.
    /// </summary>
    public IXLColumn AddVerticalPageBreak()
    {
        this.Worksheet.PageSetup.AddVerticalPageBreak(this.ColumnNumber());
        return this;
    }

    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents) =>
        this.Column(
            (this as IXLRangeBase).FirstCellUsed(options),
            (this as IXLRangeBase).LastCellUsed(options)
        );

    #endregion IXLColumn Members

    #region IXLFormatContainer

    /// <remarks>
    /// Format of the column or <c>null</c> for default format.
    /// </remarks>
    /// <inheritdoc cref="IXLFormatContainer.FormatValue"/>
    public XLCellFormatValue? FormatValue { get; set; }

    internal override XLCellFormat Format => XLCellFormat.ForColumn(this);

    private XLCellFormatValue GetFormat() =>
        this.FormatValue ?? this.Worksheet.Workbook.Styles.DefaultCellFormat;

    #endregion

    public override XLRange AsRange() => this.Range(1, 1, XlsxSharp.XLHelper.MaxRowNumber, 1);

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        return; // Columns are shifted by XLColumnCollection
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        //do nothing
    }

    internal void SetColumnNumber(int column) =>
        this.RangeAddress = new XLRangeAddress(
            new XLAddress(
                this.Worksheet,
                1,
                column,
                this.RangeAddress.FirstAddress.FixedRow,
                this.RangeAddress.FirstAddress.FixedColumn
            ),
            new XLAddress(
                this.Worksheet,
                XlsxSharp.XLHelper.MaxRowNumber,
                column,
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
                this.FixColumnAddress(firstPart) + ":" + this.FixColumnAddress(secondPart);
        }
        else
        {
            rangeAddressToUse = this.FixColumnAddress(rangeAddressStr);
        }

        XLRangeAddress rangeAddress = new(this.Worksheet, rangeAddressToUse);
        return this.Range(rangeAddress);
    }

    public IXLRangeColumn Range(int firstRow, int lastRow) =>
        this.Range(firstRow, 1, lastRow, 1).Column(1);

    private XLColumn ColumnShift(int columnsToShift) =>
        this.Worksheet.Column(this.ColumnNumber() + columnsToShift);

    #region XLColumn Left

    IXLColumn IXLColumn.ColumnLeft() => this.ColumnLeft();

    IXLColumn IXLColumn.ColumnLeft(int step) => this.ColumnLeft(step);

    public XLColumn ColumnLeft() => this.ColumnLeft(1);

    public XLColumn ColumnLeft(int step) => this.ColumnShift(step * -1);

    #endregion XLColumn Left

    #region XLColumn Right

    IXLColumn IXLColumn.ColumnRight() => this.ColumnRight();

    IXLColumn IXLColumn.ColumnRight(int step) => this.ColumnRight(step);

    public XLColumn ColumnRight() => this.ColumnRight(1);

    public XLColumn ColumnRight(int step) => this.ColumnShift(step);

    #endregion XLColumn Right

    public override bool IsEmpty() => this.IsEmpty(XLCellsUsedOptions.AllContents);

    public override bool IsEmpty(XLCellsUsedOptions options)
    {
        if (options.HasFlag(XLCellsUsedOptions.NormalFormats) && this.FormatValue is not null)
        {
            return false;
        }

        return base.IsEmpty(options);
    }

    public override bool IsEntireRow() => false;

    public override bool IsEntireColumn() => true;
}

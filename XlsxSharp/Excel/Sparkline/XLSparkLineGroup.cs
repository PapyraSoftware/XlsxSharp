#nullable disable
#nullable enable annotations

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ClosedXML.Parser;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLSparklineGroup : IXLSparklineGroup, ISheetListener
{
    private readonly XLWorksheet _worksheet;
    private readonly Dictionary<Point, SparklineFormula?> _sparklines = new();
    private IXLRange? _dateRange;
    private IXLSparklineStyle _style;

    #region Public Properties

    public IXLRange? DateRange
    {
        get => this._dateRange;
        set => this.SetDateRange(value);
    }

    public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

    public Boolean DisplayHidden { get; set; }

    public IXLSparklineHorizontalAxis HorizontalAxis { get; }

    public Double LineWeight { get; set; }

    public XLSparklineMarkers ShowMarkers { get; set; }

    public IXLSparklineStyle Style
    {
        get => this._style;
        set => this.SetStyle(value);
    }

    public XLSparklineType Type { get; set; }

    public IXLSparklineVerticalAxis VerticalAxis { get; }

    public IXLWorksheet Worksheet => this._worksheet;

    #endregion Public Properties

    /// <summary>
    /// A collection of sparkline locations and their formulas.
    /// </summary>
    internal IEnumerable<(Point Location, string? SourceDataFormula)> Sparklines =>
        this._sparklines.Select(static sl => (sl.Key, sl.Value?.Text));

    #region Public Constructors

    /// <summary>
    /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
    /// </summary>
    /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
    /// <param name="copyFrom">The sparkline group to copy from</param>
    /// <returns>The new sparkline group added</returns>
    public XLSparklineGroup(IXLWorksheet targetWorksheet, IXLSparklineGroup copyFrom)
        : this(targetWorksheet) => this.CopyFrom(copyFrom);

    /// <summary>
    /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
    /// </summary>
    /// <returns>The new sparkline group added</returns>
    public XLSparklineGroup(
        IXLWorksheet targetWorksheet,
        string locationAddress,
        string sourceDataAddress
    )
        : this(targetWorksheet) => this.Add(locationAddress, sourceDataAddress);

    /// <summary>
    /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
    /// </summary>
    /// <returns>The new sparkline group added</returns>
    public XLSparklineGroup(IXLCell location, IXLRange sourceData)
        : this(location.Worksheet) => this.Add(location, sourceData);

    /// <summary>
    /// Add a new sparkline group copied from an existing sparkline group to the specified worksheet
    /// </summary>
    /// <returns>The new sparkline group added</returns>
    public XLSparklineGroup(IXLRange locationRange, IXLRange sourceDataRange)
        : this(locationRange.Worksheet) => this.Add(locationRange, sourceDataRange);

    #endregion Public Constructors

    #region Public Methods

    public IEnumerable<IXLSparkline> Add(IXLRange locationRange, IXLRange sourceDataRange)
    {
        bool singleRow = locationRange.RowCount() == 1;
        bool singleColumn = locationRange.ColumnCount() == 1;
        List<IXLSparkline>? newSparklines = [];

        if (singleRow && singleColumn)
        {
            newSparklines.Add(this.Add(locationRange.FirstCell(), sourceDataRange));
        }
        else if (singleRow)
        {
            if (locationRange.ColumnCount() != sourceDataRange.ColumnCount())
            {
                throw new ArgumentException(
                    "locationRange and sourceDataRange must have the same width"
                );
            }

            for (int i = 1; i <= locationRange.ColumnCount(); i++)
            {
                newSparklines.Add(
                    this.Add(locationRange.Cell(1, i), sourceDataRange.Column(i).AsRange())
                );
            }
        }
        else if (singleColumn)
        {
            if (locationRange.RowCount() != sourceDataRange.RowCount())
            {
                throw new ArgumentException(
                    "locationRange and sourceDataRange must have the same height"
                );
            }

            for (int i = 1; i <= locationRange.RowCount(); i++)
            {
                newSparklines.Add(
                    this.Add(locationRange.Cell(i, 1), sourceDataRange.Row(i).AsRange())
                );
            }
        }
        else
        {
            throw new ArgumentException(
                "locationRange must have either a single row or a single column"
            );
        }

        return newSparklines;
    }

    /// <summary>
    /// Add a sparkline to the group.
    /// </summary>
    /// <param name="location">The cell to add sparklines to. If it already contains a sparkline
    /// it will be replaced.</param>
    /// <param name="sourceData">The range the sparkline gets data from</param>
    /// <returns>A newly created sparkline.</returns>
    public IXLSparkline Add(IXLCell location, IXLRange sourceData)
    {
        if (location.Worksheet != this._worksheet)
        {
            throw new ArgumentException(
                "The specified sparkline belongs to the different worksheet"
            );
        }

        // Keep invariant that each cell can have at most one sparkline
        this._worksheet.SparklineGroupsInternal.Remove(location);
        Point point = Point.FromCell(location);
        this.AddSparkline(point, sourceData);
        return new XLSparkline(this, point);
    }

    public IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress)
    {
        IXLRange? sourceDataRange =
            this._worksheet.Workbook.Range(sourceDataAddress)
            ?? this._worksheet.Range(sourceDataAddress);
        return this.Add(this._worksheet.Range(locationRangeAddress), sourceDataRange);
    }

    /// <summary>
    /// Copy the details from a specified sparkline group
    /// </summary>
    /// <param name="sparklineGroup">The sparkline group to copy from</param>
    public void CopyFrom(IXLSparklineGroup sparklineGroup)
    {
        if (sparklineGroup.DateRange != null)
        {
            this.DateRange =
                sparklineGroup.DateRange.Worksheet == sparklineGroup.Worksheet
                    ? this._worksheet.Range(sparklineGroup.DateRange.RangeAddress.ToString())
                    : sparklineGroup.DateRange;
        }

        this.DisplayEmptyCellsAs = sparklineGroup.DisplayEmptyCellsAs;
        this.DisplayHidden = sparklineGroup.DisplayHidden;
        this.LineWeight = sparklineGroup.LineWeight;
        this.ShowMarkers = sparklineGroup.ShowMarkers;
        this.Type = sparklineGroup.Type;

        XLSparklineStyle.Copy(sparklineGroup.Style, this.Style);
        XLSparklineHorizontalAxis.Copy(sparklineGroup.HorizontalAxis, this.HorizontalAxis);
        XLSparklineVerticalAxis.Copy(sparklineGroup.VerticalAxis, this.VerticalAxis);
    }

    /// <inheritdoc cref="IXLSparklineGroup.CopyTo(IXLWorksheet)"/>
    IXLSparklineGroup IXLSparklineGroup.CopyTo(IXLWorksheet targetSheet) =>
        this.CopyTo((XLWorksheet)targetSheet);

    internal XLSparklineGroup CopyTo(XLWorksheet targetSheet)
    {
        if (targetSheet == this._worksheet)
        {
            throw new InvalidOperationException(
                "Cannot copy the sparkline group to the same worksheet it belongs to"
            );
        }

        XLSparklineGroup? groupCopy = new(targetSheet, this);
        targetSheet.SparklineGroupsInternal.Add(groupCopy);
        foreach ((Point sparklineLocation, SparklineFormula? sourceData) in this._sparklines)
        {
            SparklineFormula? copiedSourceData = sourceData?.CopyFromTo(
                this._worksheet,
                targetSheet
            );
            groupCopy._sparklines.Add(sparklineLocation, copiedSourceData);
        }

        return groupCopy;
    }

    public IEnumerator<XLSparkline> GetEnumerator()
    {
        foreach (Point sparklinePoint in this._sparklines.Keys)
        {
            yield return new XLSparkline(this, sparklinePoint);
        }
    }

    IEnumerator<IXLSparkline> IEnumerable<IXLSparkline>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLSparkline? GetSparkline(IXLCell cell)
    {
        if (cell.Worksheet != this._worksheet)
        {
            return null;
        }

        Point location = Point.FromCell(cell);
        if (!this._sparklines.ContainsKey(location))
        {
            return null;
        }

        return new XLSparkline(this, location);
    }

    public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange)
    {
        if (searchRange.Worksheet != this._worksheet)
        {
            yield break;
        }

        Area searchArea = Area.FromRangeAddress(searchRange.RangeAddress);
        foreach (Point location in this._sparklines.Keys.Where(searchArea.Contains))
        {
            yield return new XLSparkline(this, location);
        }
    }

    /// <summary>
    /// Remove all sparklines in the specified cell from this group
    /// </summary>
    /// <param name="cell">The cell to remove sparklines from</param>
    public void Remove(IXLCell cell)
    {
        if (cell.Worksheet != this._worksheet)
        {
            return;
        }

        this.Remove(Point.FromCell(cell));
    }

    /// <summary>
    /// Remove the sparkline from this group
    /// </summary>
    /// <param name="sparkline"></param>
    public void Remove(IXLSparkline sparkline) => this.Remove(sparkline.Location);

    /// <summary>
    /// Remove all sparklines from this group
    /// </summary>
    public void RemoveAll() => this._sparklines.Clear();

    public IXLSparklineGroup SetDateRange(IXLRange value)
    {
        if (value != null)
        {
            if (value.RowCount() != 1 && value.ColumnCount() != 1)
            {
                throw new ArgumentException(
                    "The date range must be either one row high or one column wide"
                );
            }
        }

        this._dateRange = value;
        return this;
    }

    public IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues displayEmptyCellsAs)
    {
        this.DisplayEmptyCellsAs = displayEmptyCellsAs;
        return this;
    }

    public IXLSparklineGroup SetDisplayHidden(Boolean displayHidden)
    {
        this.DisplayHidden = displayHidden;
        return this;
    }

    public IXLSparklineGroup SetLineWeight(Double lineWeight)
    {
        this.LineWeight = lineWeight;
        return this;
    }

    public IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value)
    {
        this.ShowMarkers = value;
        return this;
    }

    public IXLSparklineGroup SetStyle(IXLSparklineStyle value)
    {
        this._style = value ?? throw new ArgumentNullException(nameof(value));
        return this;
    }

    public IXLSparklineGroup SetType(XLSparklineType type)
    {
        this.Type = type;
        return this;
    }

    #endregion Public Methods

    /// <summary>
    /// Set sparkline at the location to the specified formula.
    /// </summary>
    internal void SetSparkline(Point location, string? sourceDataFormula) =>
        this._sparklines[location] = !string.IsNullOrWhiteSpace(sourceDataFormula)
            ? new SparklineFormula(sourceDataFormula)
            : null;

    internal void Remove(Point location) => this._sparklines.Remove(location);

    internal void MoveSparkline(Point originalLocation, Point sparklineDestination)
    {
        if (!this._sparklines.TryGetValue(originalLocation, out SparklineFormula? sourceData))
        {
            throw new InvalidOperationException(
                $"No sparkline at the source cell {originalLocation}."
            );
        }

        // Target can contain sparkline from different group, ensure invariant that only one sparkline per cell
        this._worksheet.SparklineGroupsInternal.Remove(sparklineDestination);
        this._sparklines.Remove(originalLocation);
        this._sparklines[sparklineDestination] = sourceData;
    }

    internal bool TryGetSparkline(Point location, [NotNullWhen(true)] out XLSparkline? sparkline)
    {
        if (!this._sparklines.ContainsKey(location))
        {
            sparkline = null;
            return false;
        }

        sparkline = new XLSparkline(this, location);
        return true;
    }

    internal IXLRange? GetSparklineSourceData(Point sparklineLocation)
    {
        if (!this._sparklines.TryGetValue(sparklineLocation, out SparklineFormula? sourceData))
        {
            throw new InvalidOperationException(
                $"No sparkline at the source cell {sparklineLocation}."
            );
        }

        // Sparkline formula is always specified with a sheet (or a global name), it doesn't need current worksheet.
        return sourceData is not null
            ? this._worksheet.Workbook.Range(sourceData.Value.Text)
            : null;
    }

    internal void SetSparklineSourceData(Point sparklineLocation, IXLRange? sourceDataRange)
    {
        if (!this._sparklines.Remove(sparklineLocation))
        {
            throw new InvalidOperationException(
                $"No sparkline at the source cell {sparklineLocation}."
            );
        }

        this.AddSparkline(sparklineLocation, sourceDataRange);
    }

    internal IXLCell GetLocation(Point sparklineLocation)
    {
        if (!this._sparklines.ContainsKey(sparklineLocation))
        {
            throw new InvalidOperationException(
                $"No sparkline at the source cell {sparklineLocation}."
            );
        }

        return this._worksheet.Cell(sparklineLocation);
    }

    private void AddSparkline(Point location, IXLRange? sourceData)
    {
        if (sourceData is not null && sourceData.Worksheet.Workbook != this._worksheet.Workbook)
        {
            throw new ArgumentException("Range is from different workbook.");
        }

        if (sourceData is not null && sourceData.RowCount() != 1 && sourceData.ColumnCount() != 1)
        {
            throw new ArgumentException(
                "SourceData range must have either a single row or a single column"
            );
        }

        this._sparklines.Add(location, SparklineFormula.From(sourceData));
    }

    #region Private Constructors

    /// <summary>
    /// Add a new sparkline group to the specified worksheet
    /// </summary>
    /// <param name="targetWorksheet">The worksheet the sparkline group is being added to</param>
    /// <returns>The new sparkline group added</returns>
    internal XLSparklineGroup(IXLWorksheet targetWorksheet)
    {
        this._worksheet =
            targetWorksheet as XLWorksheet
            ?? throw new ArgumentNullException(nameof(targetWorksheet));
        this.HorizontalAxis = new XLSparklineHorizontalAxis(this);
        this.VerticalAxis = new XLSparklineVerticalAxis(this);
        this.HorizontalAxis.Color = XLColor.Black;
        this.Style = XLSparklineTheme.Default;
        this.LineWeight = 0.75d;
    }

    #endregion Private Constructors

    // TODO: Sparklines locations should use ST_Sqref semantic for shifting, despite constraint "This sqref element MUST contain exactly one ref element". The code assumes it just shifts individual locations points.
    #region ISheetListner

    void ISheetListener.OnInsertAreaAndShiftDown(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea insertedBookArea = new(sheet.Name, insertedArea);
        this.ShiftLocation(
            insertedBookArea,
            static (location, insertedArea) =>
            {
                if (!location.InRangeOrBelow(insertedArea))
                {
                    return location;
                }

                int shiftedRow = location.Row + insertedArea.Height;
                if (shiftedRow <= XlsxSharp.XLHelper.MaxRowNumber)
                {
                    return new Point(shiftedRow, location.Column);
                }

                return null;
            }
        );

        ReferenceShiftOnInsertRefModVisitor? refMod = new(insertedBookArea, true);
        this.AdjustSourceData(refMod);
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XLWorksheet sheet, Area insertedArea)
    {
        SheetArea insertedBookArea = new(sheet.Name, insertedArea);
        this.ShiftLocation(
            insertedBookArea,
            static (location, insertedArea) =>
            {
                if (!location.InRangeOrToRight(insertedArea))
                {
                    return location;
                }

                int shiftedColumn = location.Column + insertedArea.Width;
                if (shiftedColumn <= XlsxSharp.XLHelper.MaxColumnNumber)
                {
                    return new Point(location.Row, shiftedColumn);
                }

                return null;
            }
        );

        ReferenceShiftOnInsertRefModVisitor? refMod = new(insertedBookArea, false);
        this.AdjustSourceData(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deletedBookArea = new(sheet.Name, deletedArea);
        this.ShiftLocation(
            deletedBookArea,
            static (location, deletedArea) =>
            {
                if (!location.InRangeOrToRight(deletedArea))
                {
                    return location;
                }

                int shiftedColumn = location.Column - deletedArea.Width;
                if (shiftedColumn >= XlsxSharp.XLHelper.MinColumnNumber)
                {
                    return new Point(location.Row, shiftedColumn);
                }

                return null;
            }
        );

        ReferenceShiftOnDeleteRefModVisitor? refMod = new(
            deletedBookArea,
            XLShiftDeletedCells.ShiftCellsLeft
        );
        this.AdjustSourceData(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
    {
        SheetArea deletedBookArea = new(sheet.Name, deletedArea);
        this.ShiftLocation(
            deletedBookArea,
            static (location, deletedArea) =>
            {
                if (!location.InRangeOrBelow(deletedArea))
                {
                    return location;
                }

                int shiftedRow = location.Row - deletedArea.Height;
                if (shiftedRow is >= XlsxSharp.XLHelper.MinRowNumber)
                {
                    return new Point(shiftedRow, location.Column);
                }

                return null;
            }
        );

        ReferenceShiftOnDeleteRefModVisitor? refMod = new(
            deletedBookArea,
            XLShiftDeletedCells.ShiftCellsUp
        );
        this.AdjustSourceData(refMod);
    }

    private void ShiftLocation(SheetArea shiftedRange, Func<Point, Area, Point?> shiftLocation)
    {
        // If shift was on another worksheet, there is no way to affect sparklines for this worksheet of this group
        if (!XlsxSharp.XLHelper.SheetComparer.Equals(shiftedRange.Name, this._worksheet.Name))
        {
            return;
        }

        Dictionary<Point, SparklineFormula?>? sparklinesCopy = new(this._sparklines);

        // Clear to avoid problems during shifting (e.g. A1 and A2 have sparklines, A1 is
        // shifted to A2, but A2 hasn't yet been shifted). Just reinsert everything.
        this._sparklines.Clear();
        foreach ((Point originalLocation, SparklineFormula? sourceData) in sparklinesCopy)
        {
            Point? shiftedLocation = shiftLocation(originalLocation, shiftedRange.Area);
            if (shiftedLocation is not null)
            {
                this._sparklines.Add(shiftedLocation.Value, sourceData);
            }
        }
    }

    private void AdjustSourceData(CopyVisitor refMod)
    {
        // Can't modify dictionary while iterating over it, make a copy.
        List<Point>? locationsCopy = [.. this._sparklines.Keys];
        foreach (Point location in locationsCopy)
        {
            SparklineFormula? originalSourceData = this._sparklines[location];
            if (originalSourceData is not null)
            {
                string? shiftedSourceData = FormulaConverter.ModifyA1(
                    originalSourceData.Value.Text,
                    this._worksheet.Name,
                    location.Row,
                    location.Column,
                    refMod
                );
                this._sparklines[location] = new SparklineFormula(shiftedSourceData);
            }
        }
    }

    #endregion

    /// <summary>
    /// The source data area referenced by a sparkline. The grammar is should rather limited:
    /// <c>sparkline-formula = single-sheet-area / [single-sheet-prefix / book-prefix] name</c>.
    /// Additionally, if a single-sheet - area is specified, that single-sheet-area MUST contain cells from either
    /// a single row or a single column. In reality, it can be more encompassing (e.g. <c>'[1]Contract Tail YLT'!B46:E46</c>).
    /// </summary>
    /// <param name="Text">Text of the formula.</param>
    private readonly record struct SparklineFormula(string Text)
    {
        /// <summary>
        /// Factory method to create a formula from a reference with a sheet from the range.
        /// </summary>
        [return: NotNullIfNotNull(nameof(range))]
        internal static SparklineFormula? From(IXLRange? range)
        {
            if (range is null)
            {
                return null;
            }

            string? formula = range.RangeAddress.ToStringRelative(true);
            return new SparklineFormula(formula);
        }

        /// <summary>
        /// A factory method used for copying worksheets. If formula is a sheet reference/name,
        /// move the formula of <paramref name="sourceSheet"/> to the <paramref name="targetSheet"/>.
        /// Otherwise, return the original formula.
        /// </summary>
        internal SparklineFormula CopyFromTo(XLWorksheet sourceSheet, XLWorksheet targetSheet)
        {
            // If formula is single-sheet-area, i.e. `single-sheet-prefix A1-area`
            if (
                ReferenceParser.TryParseSheetA1(
                    this.Text,
                    out string? formulaSheetName,
                    out ReferenceArea reference
                ) && XlsxSharp.XLHelper.SheetComparer.Equals(formulaSheetName, sourceSheet.Name)
            )
            {
                string? copiedReference = reference.GetDisplayStringA1(targetSheet.Name);
                return new SparklineFormula(copiedReference);
            }

            // If formula is a `single-sheet-prefix name`
            if (
                ReferenceParser.TryParseSheetName(
                    this.Text,
                    out formulaSheetName,
                    out string? definedName
                ) && XlsxSharp.XLHelper.SheetComparer.Equals(formulaSheetName, sourceSheet.Name)
            )
            {
                string? copiedName = definedName.GetSheetDefinedName(targetSheet.Name);
                return new SparklineFormula(copiedName);
            }

            // Either just name or from different workbook
            return this;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XlsxSharp.Excel.Caching;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.Charts;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Index;
using XlsxSharp.Excel.InsertData;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.Protection.XLProtectionAlgorithm;

namespace XlsxSharp.Excel;

internal class XLWorksheet : XLRangeBase, IXLWorksheet, IXLFormatContainer
{
    #region Fields

    private readonly Dictionary<int, int> _columnOutlineCount = new();
    private readonly Dictionary<int, int> _rowOutlineCount = new();
    private readonly XLRangeFactory _rangeFactory;
    private readonly XLRangeRepository _rangeRepository;
    private readonly List<IXLRangeIndex> _rangeIndices;
    private readonly XLRanges _selectedRanges;

    internal int ZOrder = 1;
    private string _name;
    internal int _position;

    private double _rowHeight;
    private bool _tabActive;
    private XLSheetProtection _protection;

    /// <summary>
    /// Fake address to be used everywhere the invalid address is needed.
    /// </summary>
    internal readonly XLAddress InvalidAddress;

    #endregion Fields

    #region Constructor

    public XLWorksheet(string sheetName, XLWorkbook workbook, uint sheetId)
        : base(
            new XLRangeAddress(
                new XLAddress(
                    null,
                    XlsxSharp.XLHelper.MinRowNumber,
                    XlsxSharp.XLHelper.MinColumnNumber,
                    false,
                    false
                ),
                new XLAddress(
                    null,
                    XlsxSharp.XLHelper.MaxRowNumber,
                    XlsxSharp.XLHelper.MaxColumnNumber,
                    false,
                    false
                )
            )
        )
    {
        this.Workbook = workbook;
        this.SheetId = sheetId;
        this.InvalidAddress = new XLAddress(this, 0, 0, false, false);

        XLAddress firstAddress = new(
            this,
            this.RangeAddress.FirstAddress.RowNumber,
            this.RangeAddress.FirstAddress.ColumnNumber,
            this.RangeAddress.FirstAddress.FixedRow,
            this.RangeAddress.FirstAddress.FixedColumn
        );
        XLAddress lastAddress = new(
            this,
            this.RangeAddress.LastAddress.RowNumber,
            this.RangeAddress.LastAddress.ColumnNumber,
            this.RangeAddress.LastAddress.FixedRow,
            this.RangeAddress.LastAddress.FixedColumn
        );
        this.RangeAddress = new XLRangeAddress(firstAddress, lastAddress);
        this._rangeFactory = new XLRangeFactory(this);
        this._rangeRepository = new XLRangeRepository(workbook, this._rangeFactory.Create);
        this._rangeIndices = [];

        this.Pictures = new XLPictures(this);
        this.DefinedNames = new XLDefinedNames(this);
        this.SheetView = new XLSheetView(this);
        this.Tables = [];
        this.Hyperlinks = new XLHyperlinks(this);
        this.DataValidations = new XLDataValidations(this);
        this.PivotTables = new XLPivotTables(this);
        this._protection = new XLSheetProtection(DefaultProtectionAlgorithm);
        this.AutoFilter = new XLAutoFilter();
        this.ConditionalFormats = new XLConditionalFormats(this);
        this.SparklineGroupsInternal = new XLSparklineGroups(this);
        this.Internals = new XLWorksheetInternals
        {
            CellsCollection = new XLCellsCollection(this),
            ColumnsCollection = new XLColumnsCollection(),
            RowsCollection = new XLRowsCollection(),
            MergedRanges = new XLRanges(this),
        };
        this.PageSetup = new XLPageSetup((XLPageSetup)workbook.PageOptions, this);
        this.Outline = new XLOutline(workbook.Outline);
        this._columnWidth = workbook.ColumnWidth;
        this._rowHeight = workbook.RowHeight;
        this.RowHeightChanged =
            Math.Abs(workbook.RowHeight - XLWorkbook.DefaultRowHeight) > XlsxSharp.XLHelper.Epsilon;

        XlsxSharp.XLHelper.ValidateSheetName(sheetName);
        this._name = sheetName;
        this.Charts = new XLCharts();
        this.ShowFormulas = workbook.ShowFormulas;
        this.ShowGridLines = workbook.ShowGridLines;
        this.ShowOutlineSymbols = workbook.ShowOutlineSymbols;
        this.ShowRowColHeaders = workbook.ShowRowColHeaders;
        this.ShowRuler = workbook.ShowRuler;
        this.ShowWhiteSpace = workbook.ShowWhiteSpace;
        this.ShowZeros = workbook.ShowZeros;
        this.RightToLeft = workbook.RightToLeft;
        this.TabColor = XLColor.Automatic;
        this._selectedRanges = new XLRanges(this);

        this.Author = workbook.Author;
    }

    #endregion Constructor

    internal SheetArea Area => new(this.Name, Excel.Area.Full);

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXLDefinedNames IXLWorksheet.NamedRanges => this.DefinedNames;

    IXLDefinedNames IXLWorksheet.DefinedNames => this.DefinedNames;

    internal XLDefinedNames DefinedNames { get; }

    public override XLRangeType RangeType => XLRangeType.Worksheet;

    /// <summary>
    /// Reference to a VML that contains notes, forms controls and so on. All such things are generally unified into
    /// a single legacy VML file, set during load/save.
    /// </summary>
    public string? LegacyDrawingId;

    private double _columnWidth;

    public XLWorksheetInternals Internals { get; }

    internal XLSparklineGroups SparklineGroupsInternal { get; }

    public XLRangeFactory RangeFactory => this._rangeFactory;

    internal bool RowHeightChanged { get; set; }

    internal bool ColumnWidthChanged { get; set; }

    /// <summary>
    /// <para>
    /// The id of a sheet that is unique across the workbook, kept across load/save.
    /// The ids of sheets are not reused. That is important for referencing the sheet
    /// range/point through sheetId. If sheetIds were reused, references would refer
    /// to the wrong sheet after the original sheetId was reused. Excel also doesn't
    /// reuse sheetIds.
    /// </para>
    /// <para>
    /// Referencing sheet through non-reused sheetIds means that reference can survive
    /// sheet renaming without any changes. Always &gt; 0 (Excel will crash on 0).
    /// </para>
    /// </summary>
    internal uint SheetId { get; set; }

    /// <summary>
    /// A cached <c>r:id</c> of the sheet from the file. If the sheet is a new one (not
    /// yet saved), the value is null until workbook is saved. Use <see cref="SheetId"/>
    /// instead is possible. Mostly for removing deleted sheet parts during save.
    /// </summary>
    internal string? RelId { get; set; }

    public XLDataValidations DataValidations { get; private set; }

    public IXLCharts Charts { get; private set; }

    public XLSheetProtection Protection
    {
        get => this._protection;
        set => this._protection = value.Clone().CastTo<XLSheetProtection>();
    }

    public XLAutoFilter AutoFilter { get; private set; }

    public bool IsDeleted { get; private set; }

    #region IXLFormatContainer

    /// <remarks>
    /// Format of a worksheet or <c>null</c> if worksheet has no format. In OOXML, worksheet
    /// doesn't actually has a format by itself, so this is mostly virtual. If all columns have
    /// format during the load, it is set to the format of the last column in a sheet (XFD).
    /// </remarks>
    /// <inheritdoc cref="IXLFormatContainer.FormatValue"/>
    public XLCellFormatValue? FormatValue { get; set; }

    internal override XLCellFormat Format => XLCellFormat.ForWorksheet(this);

    internal XLCellFormatValue GetFormat() =>
        this.FormatValue ?? this.Workbook.Styles.DefaultCellFormat;

    #endregion

    #region IXLWorksheet Members

    public XLWorkbook Workbook { get; }

    public double ColumnWidth
    {
        get => this._columnWidth;
        set
        {
            this.ColumnWidthChanged = true;
            this._columnWidth = value;
        }
    }

    public double RowHeight
    {
        get => this._rowHeight;
        set
        {
            this.RowHeightChanged = true;
            this._rowHeight = value;
        }
    }

    public string Name
    {
        get => this._name;
        set
        {
            if (this._name == value)
            {
                return;
            }

            XlsxSharp.XLHelper.ValidateSheetName(value);

            this.Workbook.WorksheetsInternal.Rename(this._name, value);
            this._name = value;
        }
    }

    public int Position
    {
        get => this._position;
        set
        {
            if (
                value
                > this.Workbook.WorksheetsInternal.Count + this.Workbook.UnsupportedSheets.Count + 1
            )
            {
                throw new ArgumentOutOfRangeException(
                    nameof(value),
                    "Index must be equal or less than the number of worksheets + 1."
                );
            }

            if (value < this._position)
            {
                this.Workbook.WorksheetsInternal.Where<XLWorksheet>(w =>
                        w.Position >= value && w.Position < this._position
                    )
                    .ForEach(w => w._position += 1);
            }

            if (value > this._position)
            {
                this.Workbook.WorksheetsInternal.Where<XLWorksheet>(w =>
                        w.Position <= value && w.Position > this._position
                    )
                    .ForEach(w => (w)._position -= 1);
            }

            this._position = value;
        }
    }

    public IXLPageSetup PageSetup { get; private set; }

    public IXLOutline Outline { get; private set; }

    IXLRow? IXLWorksheet.FirstRowUsed() => this.FirstRowUsed();

    IXLRow? IXLWorksheet.FirstRowUsed(XLCellsUsedOptions options) => this.FirstRowUsed(options);

    IXLRow? IXLWorksheet.LastRowUsed() => this.LastRowUsed();

    IXLRow? IXLWorksheet.LastRowUsed(XLCellsUsedOptions options) => this.LastRowUsed(options);

    IXLColumn IXLWorksheet.LastColumn() => this.LastColumn();

    IXLColumn IXLWorksheet.FirstColumn() => this.FirstColumn();

    IXLRow IXLWorksheet.FirstRow() => this.FirstRow();

    IXLRow IXLWorksheet.LastRow() => this.LastRow();

    IXLColumn? IXLWorksheet.FirstColumnUsed() => this.FirstColumnUsed();

    IXLColumn? IXLWorksheet.FirstColumnUsed(XLCellsUsedOptions options) =>
        this.FirstColumnUsed(options);

    IXLColumn? IXLWorksheet.LastColumnUsed() => this.LastColumnUsed();

    IXLColumn? IXLWorksheet.LastColumnUsed(XLCellsUsedOptions options) =>
        this.LastColumnUsed(options);

    public IXLColumns Columns()
    {
        HashSet<int> columnMap = [];

        columnMap.UnionWith(this.Internals.CellsCollection.ColumnsUsedKeys);
        columnMap.UnionWith(this.Internals.ColumnsCollection.Keys);

        XLColumns retVal = new(this.Workbook, this, this, columnMap.Select(this.Column));
        return retVal;
    }

    public IXLColumns Columns(string columns)
    {
        XLColumns retVal = new(this.Workbook, null, this);
        string[] columnPairs = columns.Split(',');
        foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
        {
            string firstColumn;
            string lastColumn;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] columnRange = XlsxSharp.XLHelper.SplitRange(tPair);
                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = tPair;
                lastColumn = tPair;
            }

            if (int.TryParse(firstColumn, out int tmp))
            {
                foreach (
                    IXLColumn col in this.Columns(int.Parse(firstColumn), int.Parse(lastColumn))
                )
                {
                    retVal.Add((XLColumn)col);
                }
            }
            else
            {
                foreach (IXLColumn col in this.Columns(firstColumn, lastColumn))
                {
                    retVal.Add((XLColumn)col);
                }
            }
        }
        return retVal;
    }

    public IXLColumns Columns(string firstColumn, string lastColumn) =>
        this.Columns(
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(firstColumn),
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(lastColumn)
        );

    public IXLColumns Columns(int firstColumn, int lastColumn)
    {
        XLColumns retVal = new(this.Workbook, null, this);

        for (int co = firstColumn; co <= lastColumn; co++)
        {
            retVal.Add(this.Column(co));
        }

        return retVal;
    }

    public IXLRows Rows()
    {
        HashSet<int> rowMap = [];

        rowMap.UnionWith(this.Internals.CellsCollection.RowsUsedKeys);
        rowMap.UnionWith(this.Internals.RowsCollection.Keys);

        XLRows retVal = new(this.Workbook, this, this, rowMap.Select(this.Row));
        return retVal;
    }

    public IXLRows Rows(string rows)
    {
        XLRows retVal = new(this.Workbook, null, this);
        string[] rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                string[] rowRange = XlsxSharp.XLHelper.SplitRange(tPair);
                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }

            this.Rows(int.Parse(firstRow), int.Parse(lastRow))
                .ForEach(row => retVal.Add((XLRow)row));
        }
        return retVal;
    }

    public IXLRows Rows(int firstRow, int lastRow)
    {
        XLRows retVal = new(this.Workbook, null, this);

        for (int ro = firstRow; ro <= lastRow; ro++)
        {
            retVal.Add(this.Row(ro));
        }

        return retVal;
    }

    IXLRow IXLWorksheet.Row(int row) => this.Row(row);

    IXLColumn IXLWorksheet.Column(int column) => this.Column(column);

    IXLColumn IXLWorksheet.Column(string column) => this.Column(column);

    IXLCell IXLWorksheet.Cell(int row, int column) => this.Cell(row, column);

    IXLCell IXLWorksheet.Cell(string cellAddressInRange) =>
        this.Cell(cellAddressInRange)
        ?? throw new ArgumentException(
            $"'{cellAddressInRange}' is not A1 address or workbook named range."
        );

    IXLCell IXLWorksheet.Cell(int row, string column) => this.Cell(row, column);

    IXLCell IXLWorksheet.Cell(IXLAddress cellAddressInRange) => this.Cell(cellAddressInRange);

    IXLRange IXLWorksheet.Range(IXLRangeAddress rangeAddress) => this.Range(rangeAddress);

    IXLRange IXLWorksheet.Range(string rangeAddress) =>
        this.Range(rangeAddress)
        ?? throw new ArgumentException($"'{rangeAddress}' is not A1 address or named range.");

    IXLRange IXLWorksheet.Range(IXLCell firstCell, IXLCell lastCell) =>
        this.Range(firstCell, lastCell);

    IXLRange IXLWorksheet.Range(string firstCellAddress, string lastCellAddress) =>
        this.Range(firstCellAddress, lastCellAddress);

    IXLRange IXLWorksheet.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress) =>
        this.Range(firstCellAddress, lastCellAddress);

    IXLRange IXLWorksheet.Range(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    ) => this.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);

    IXLRanges IXLWorksheet.Ranges(string ranges) => this.Ranges(ranges);

    public IXLWorksheet CollapseRows()
    {
        Enumerable.Range(1, 8).ForEach(i => this.CollapseRows(i));
        return this;
    }

    public IXLWorksheet CollapseColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => this.CollapseColumns(i));
        return this;
    }

    public IXLWorksheet ExpandRows()
    {
        Enumerable.Range(1, 8).ForEach(i => this.ExpandRows(i));
        return this;
    }

    public IXLWorksheet ExpandColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => this.ExpandColumns(i));
        return this;
    }

    public IXLWorksheet CollapseRows(int outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
        {
            throw new ArgumentOutOfRangeException(
                "outlineLevel",
                "Outline level must be between 1 and 8."
            );
        }

        this.Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel)
            .ForEach(r => r.Collapse());
        return this;
    }

    public IXLWorksheet CollapseColumns(int outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
        {
            throw new ArgumentOutOfRangeException(
                "outlineLevel",
                "Outline level must be between 1 and 8."
            );
        }

        this.Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel)
            .ForEach(c => c.Collapse());
        return this;
    }

    public IXLWorksheet ExpandRows(int outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
        {
            throw new ArgumentOutOfRangeException(
                "outlineLevel",
                "Outline level must be between 1 and 8."
            );
        }

        this.Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel)
            .ForEach(r => r.Expand());
        return this;
    }

    public IXLWorksheet ExpandColumns(int outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
        {
            throw new ArgumentOutOfRangeException(
                "outlineLevel",
                "Outline level must be between 1 and 8."
            );
        }

        this.Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel)
            .ForEach(c => c.Expand());
        return this;
    }

    public void Delete()
    {
        this.IsDeleted = true;
        this.Workbook.DefinedNamesInternal.OnWorksheetDeleted(this.Name);
        this.Workbook.NotifyWorksheetDeleting(this);
        this.Workbook.WorksheetsInternal.Delete(this.Name);
    }

    [Obsolete($"Used {nameof(DefinedName)} instead.")]
    IXLDefinedName IXLWorksheet.NamedRange(string name) => this.DefinedName(name);

    IXLDefinedName IXLWorksheet.DefinedName(string name) => this.DefinedName(name);

    internal XLDefinedName DefinedName(string name) => this.DefinedNames.DefinedName(name);

    IXLSheetView IXLWorksheet.SheetView => this.SheetView;

    public XLSheetView SheetView { get; private set; }

    IXLTables IXLWorksheet.Tables => this.Tables;

    internal XLTables Tables { get; }

    public IXLTable Table(int index) => this.Tables.Table(index);

    public IXLTable Table(string name) => this.Tables.Table(name);

    public IXLWorksheet CopyTo(string newSheetName) =>
        this.CopyTo(this.Workbook, newSheetName, this.Workbook.WorksheetsInternal.Count + 1);

    public IXLWorksheet CopyTo(string newSheetName, int position) =>
        this.CopyTo(this.Workbook, newSheetName, position);

    public IXLWorksheet CopyTo(XLWorkbook workbook) =>
        this.CopyTo(workbook, this.Name, workbook.WorksheetsInternal.Count + 1);

    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName) =>
        this.CopyTo(workbook, newSheetName, workbook.WorksheetsInternal.Count + 1);

    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName, int position)
    {
        if (this.IsDeleted)
        {
            throw new InvalidOperationException(
                $"`{this.Name}` has been deleted and cannot be copied."
            );
        }

        XLWorksheet targetSheet = (XLWorksheet)
            workbook.WorksheetsInternal.Add(newSheetName, position);
        this.Internals.ColumnsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Column(kp.Key)));
        this.Internals.RowsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Row(kp.Key)));
        this.Internals.CellsCollection.GetCells()
            .ForEach(c =>
                targetSheet
                    .Cell(c.Address)
                    .CopyFrom(c, XLCellCopyOptions.Values | XLCellCopyOptions.Styles)
            );
        foreach (XLDataValidation dataValidation in this.DataValidations)
        {
            targetSheet.DataValidations.CopyFrom(dataValidation);
        }

        targetSheet.Visibility = this.Visibility;
        targetSheet.ColumnWidth = this.ColumnWidth;
        targetSheet.ColumnWidthChanged = this.ColumnWidthChanged;
        targetSheet.RowHeight = this.RowHeight;
        targetSheet.RowHeightChanged = this.RowHeightChanged;
        if (this.FormatValue is not null)
        {
            targetSheet.FormatValue = workbook.Styles.GetRegisteredCellFormat(this.FormatValue);
        }

        targetSheet.PageSetup = new XLPageSetup((XLPageSetup)this.PageSetup, targetSheet);
        ((XLHeaderFooter)targetSheet.PageSetup.Header).Changed = true;
        ((XLHeaderFooter)targetSheet.PageSetup.Footer).Changed = true;
        targetSheet.Outline = new XLOutline(this.Outline);
        targetSheet.SheetView = new XLSheetView(targetSheet, this.SheetView);
        targetSheet.SelectedRanges.RemoveAll();

        foreach (XLPicture picture in this.Pictures)
        {
            picture.CopyTo(targetSheet);
        }

        this.Tables.ForEach<XLTable>(t => t.CopyTo(targetSheet, false));
        this.DefinedNames.ForEach<XLDefinedName>(nr => nr.CopyTo(targetSheet)); // Names must modify table references, so keep the order.
        this.PivotTables.ForEach<XLPivotTable>(pt =>
            pt.CopyTo(
                targetSheet.Cell(pt.TargetCell.Address.CastTo<XLAddress>().WithoutWorksheet())
            )
        );
        foreach (XLConditionalFormat cf in this.ConditionalFormats)
        {
            cf.CopyTo(targetSheet);
        }

        // Sparklines were already copied during copy of columns, rows and cells, but piecemeal (e.g. multi-cell
        // sparkline group could be split into group-per-cell). Since this is a copy of whole sheet, just remove
        // the piecemeal copy and copy it propertly.
        targetSheet.SparklineGroupsInternal.RemoveAll();
        this.SparklineGroups.CopyTo(targetSheet);
        this.MergedRanges.ForEach(mr =>
            targetSheet.Range(((XLRangeAddress)mr.RangeAddress).WithoutWorksheet()).Merge()
        );
        this.SelectedRanges.ForEach(sr =>
            targetSheet.SelectedRanges.Add(
                targetSheet.Range(((XLRangeAddress)sr.RangeAddress).WithoutWorksheet())
            )
        );

        if (this.AutoFilter.IsEnabled)
        {
            XLRange range = targetSheet.Range(
                ((XLRangeAddress)this.AutoFilter.Range.RangeAddress).WithoutWorksheet()
            );
            range.SetAutoFilter();
        }

        return targetSheet;
    }

    internal XLHyperlinks Hyperlinks { get; }

    IXLHyperlinks IXLWorksheet.Hyperlinks => this.Hyperlinks;

    IXLDataValidations IXLWorksheet.DataValidations => this.DataValidations;

    private XLWorksheetVisibility _visibility;

    public XLWorksheetVisibility Visibility
    {
        get => this._visibility;
        set
        {
            if (value != XLWorksheetVisibility.Visible)
            {
                this.TabSelected = false;
            }

            this._visibility = value;
        }
    }

    public IXLWorksheet Hide()
    {
        this.Visibility = XLWorksheetVisibility.Hidden;
        return this;
    }

    public IXLWorksheet Unhide()
    {
        this.Visibility = XLWorksheetVisibility.Visible;
        return this;
    }

    IXLSheetProtection IXLProtectable<IXLSheetProtection, XLSheetProtectionElements>.Protection
    {
        get => this.Protection;
        set => this.Protection = (XLSheetProtection)value;
    }

    public IXLSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm) =>
        this.Protection.Protect(algorithm);

    public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements) =>
        this.Protection.Protect(allowedElements);

    public IXLSheetProtection Protect(
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    ) => this.Protection.Protect(algorithm, allowedElements);

    public IXLSheetProtection Protect(
        string password,
        Algorithm algorithm = DefaultProtectionAlgorithm
    ) => this.Protection.Protect(password, algorithm);

    public IXLSheetProtection Protect(
        string password,
        Algorithm algorithm,
        XLSheetProtectionElements allowedElements
    ) => this.Protection.Protect(password, algorithm, allowedElements);

    IXLElementProtection IXLProtectable.Protect(Algorithm algorithm) => this.Protect(algorithm);

    IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm) =>
        this.Protect(password, algorithm);

    public IXLSheetProtection Unprotect() => this.Protection.Unprotect();

    public IXLSheetProtection Unprotect(string password) => this.Protection.Unprotect(password);

    IXLElementProtection IXLProtectable.Unprotect() => this.Unprotect();

    IXLElementProtection IXLProtectable.Unprotect(string password) => this.Unprotect(password);

    public new IXLRange Sort() => this.GetRangeForSort().Sort();

    public new IXLRange Sort(
        string columnsToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    ) => this.GetRangeForSort().Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks);

    public new IXLRange Sort(
        int columnToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    ) => this.GetRangeForSort().Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);

    public new IXLRange SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    ) => this.GetRangeForSort().SortLeftToRight(sortOrder, matchCase, ignoreBlanks);

    public bool ShowFormulas { get; set; }

    public bool ShowGridLines { get; set; }

    public bool ShowOutlineSymbols { get; set; }

    public bool ShowRowColHeaders { get; set; }

    public bool ShowRuler { get; set; }

    public bool ShowWhiteSpace { get; set; }

    public bool ShowZeros { get; set; }

    public IXLWorksheet SetShowFormulas()
    {
        this.ShowFormulas = true;
        return this;
    }

    public IXLWorksheet SetShowFormulas(bool value)
    {
        this.ShowFormulas = value;
        return this;
    }

    public IXLWorksheet SetShowGridLines()
    {
        this.ShowGridLines = true;
        return this;
    }

    public IXLWorksheet SetShowGridLines(bool value)
    {
        this.ShowGridLines = value;
        return this;
    }

    public IXLWorksheet SetShowOutlineSymbols()
    {
        this.ShowOutlineSymbols = true;
        return this;
    }

    public IXLWorksheet SetShowOutlineSymbols(bool value)
    {
        this.ShowOutlineSymbols = value;
        return this;
    }

    public IXLWorksheet SetShowRowColHeaders()
    {
        this.ShowRowColHeaders = true;
        return this;
    }

    public IXLWorksheet SetShowRowColHeaders(bool value)
    {
        this.ShowRowColHeaders = value;
        return this;
    }

    public IXLWorksheet SetShowRuler()
    {
        this.ShowRuler = true;
        return this;
    }

    public IXLWorksheet SetShowRuler(bool value)
    {
        this.ShowRuler = value;
        return this;
    }

    public IXLWorksheet SetShowWhiteSpace()
    {
        this.ShowWhiteSpace = true;
        return this;
    }

    public IXLWorksheet SetShowWhiteSpace(bool value)
    {
        this.ShowWhiteSpace = value;
        return this;
    }

    public IXLWorksheet SetShowZeros()
    {
        this.ShowZeros = true;
        return this;
    }

    public IXLWorksheet SetShowZeros(bool value)
    {
        this.ShowZeros = value;
        return this;
    }

    public XLColor TabColor { get; set; }

    public IXLWorksheet SetTabColor(XLColor color)
    {
        this.TabColor = color;
        return this;
    }

    public bool TabSelected { get; set; }

    public bool TabActive
    {
        get => this._tabActive;
        set
        {
            if (value && !this._tabActive)
            {
                foreach (XLWorksheet ws in this.Worksheet.Workbook.WorksheetsInternal)
                {
                    ws._tabActive = false;
                }
            }
            this._tabActive = value;
        }
    }

    public IXLWorksheet SetTabSelected()
    {
        this.TabSelected = true;
        return this;
    }

    public IXLWorksheet SetTabSelected(bool value)
    {
        this.TabSelected = value;
        return this;
    }

    public IXLWorksheet SetTabActive()
    {
        this.TabActive = true;
        return this;
    }

    public IXLWorksheet SetTabActive(bool value)
    {
        this.TabActive = value;
        return this;
    }

    IXLPivotTable IXLWorksheet.PivotTable(string name) => this.PivotTable(name);

    IXLPivotTables IXLWorksheet.PivotTables => this.PivotTables;

    public XLPivotTables PivotTables { get; }

    public bool RightToLeft { get; set; }

    public IXLWorksheet SetRightToLeft()
    {
        this.RightToLeft = true;
        return this;
    }

    public IXLWorksheet SetRightToLeft(bool value)
    {
        this.RightToLeft = value;
        return this;
    }

    public override XLRanges Ranges(string ranges)
    {
        XLRanges retVal = new(this.Workbook);
        foreach (string rangeAddressStr in ranges.Split(',').Select(s => s.Trim()))
        {
            if (rangeAddressStr.StartsWith("#REF!"))
            {
                continue;
            }
            else if (XlsxSharp.XLHelper.IsValidRangeAddress(rangeAddressStr))
            {
                retVal.Add(this.Range(new XLRangeAddress(this.Worksheet, rangeAddressStr)));
            }
            else if (
                this.DefinedNames.TryGetValue(
                    rangeAddressStr,
                    out IXLDefinedName? worksheetNamedRange
                )
            )
            {
                worksheetNamedRange.Ranges.ForEach(retVal.Add);
            }
            else if (
                this.Workbook.DefinedNames.TryGetValue(
                    rangeAddressStr,
                    out IXLDefinedName? workbookDefinedName
                )
                && workbookDefinedName.Ranges.First().Worksheet == this
            )
            {
                workbookDefinedName.Ranges.ForEach(retVal.Add);
            }
        }
        return retVal;
    }

    IXLAutoFilter IXLWorksheet.AutoFilter => this.AutoFilter;

    public IXLRows RowsUsed(
        XLCellsUsedOptions options = XLCellsUsedOptions.AllContents,
        Func<IXLRow, bool>? predicate = null
    )
    {
        XLRows rows = new(this.Workbook, worksheet: null, this);
        HashSet<int> rowsUsed = [];
        foreach (
            int rowNum in this.Internals.RowsCollection.Keys.Concat(
                this.Internals.CellsCollection.RowsUsedKeys
            )
        )
        {
            if (!rowsUsed.Add(rowNum))
            {
                continue;
            }
            XLRow row = this.Row(rowNum);
            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
            {
                rows.Add(row);
            }
        }
        return rows;
    }

    public IXLRows RowsUsed(Func<IXLRow, bool>? predicate = null) =>
        this.RowsUsed(XLCellsUsedOptions.AllContents, predicate);

    public IXLColumns ColumnsUsed(
        XLCellsUsedOptions options = XLCellsUsedOptions.AllContents,
        Func<IXLColumn, bool>? predicate = null
    )
    {
        XLColumns columns = new(this.Workbook, worksheet: null, defaultStyleSheet: this);
        HashSet<int> columnsUsed = [];
        this.Internals.ColumnsCollection.Keys.ForEach(r => columnsUsed.Add(r));
        this.Internals.CellsCollection.ColumnsUsedKeys.ForEach(r => columnsUsed.Add(r));
        foreach (int columnNum in columnsUsed)
        {
            XLColumn column = this.Column(columnNum);
            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
            {
                columns.Add(column);
            }
        }
        return columns;
    }

    public IXLColumns ColumnsUsed(Func<IXLColumn, bool>? predicate = null) =>
        this.ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);

    internal void RegisterRangeIndex(IXLRangeIndex rangeIndex) =>
        this._rangeIndices.Add(rangeIndex);

    internal void Cleanup()
    {
        this.Internals.Dispose();
        foreach (XLPicture picture in this.Pictures)
        {
            picture.Dispose();
        }

        this._rangeRepository.Clear();
        this._rangeIndices.Clear();
    }

    #endregion IXLWorksheet Members

    #region Outlines

    public void IncrementColumnOutline(int level)
    {
        if (level <= 0)
        {
            return;
        }

        if (this._columnOutlineCount.TryGetValue(level, out int value))
        {
            this._columnOutlineCount[level] = value + 1;
        }
        else
        {
            this._columnOutlineCount.Add(level, 1);
        }
    }

    public void DecrementColumnOutline(int level)
    {
        if (level <= 0)
        {
            return;
        }

        if (this._columnOutlineCount.TryGetValue(level, out int value))
        {
            if (value > 0)
            {
                this._columnOutlineCount[level] = value - 1;
            }
        }
        else
        {
            this._columnOutlineCount.Add(level, 0);
        }
    }

    public int GetMaxColumnOutline()
    {
        List<KeyValuePair<int, int>> list = [.. this._columnOutlineCount.Where(kp => kp.Value > 0)];
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }

    public void IncrementRowOutline(int level)
    {
        if (level <= 0)
        {
            return;
        }

        if (this._rowOutlineCount.TryGetValue(level, out int value))
        {
            this._rowOutlineCount[level] = value + 1;
        }
        else
        {
            this._rowOutlineCount.Add(level, 0);
        }
    }

    public void DecrementRowOutline(int level)
    {
        if (level <= 0)
        {
            return;
        }

        if (this._rowOutlineCount.TryGetValue(level, out int value))
        {
            if (value > 0)
            {
                this._rowOutlineCount[level] = level - 1;
            }
        }
        else
        {
            this._rowOutlineCount.Add(level, 0);
        }
    }

    public int GetMaxRowOutline() =>
        this._rowOutlineCount.Count == 0
            ? 0
            : this._rowOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);

    #endregion Outlines

    public XLRow? FirstRowUsed() => this.FirstRowUsed(XLCellsUsedOptions.AllContents);

    public XLRow? FirstRowUsed(XLCellsUsedOptions options)
    {
        XLRangeRow? rngRow = this.AsRange().FirstRowUsed(options);
        return rngRow != null ? this.Row(rngRow.RangeAddress.FirstAddress.RowNumber) : null;
    }

    public XLRow? LastRowUsed() => this.LastRowUsed(XLCellsUsedOptions.AllContents);

    public XLRow? LastRowUsed(XLCellsUsedOptions options)
    {
        XLRangeRow? rngRow = this.AsRange().LastRowUsed(options);
        return rngRow != null ? this.Row(rngRow.RangeAddress.LastAddress.RowNumber) : null;
    }

    public XLColumn LastColumn() => this.Column(XlsxSharp.XLHelper.MaxColumnNumber);

    public XLColumn FirstColumn() => this.Column(1);

    public XLRow FirstRow() => this.Row(1);

    public XLRow LastRow() => this.Row(XlsxSharp.XLHelper.MaxRowNumber);

    public XLColumn? FirstColumnUsed() => this.FirstColumnUsed(XLCellsUsedOptions.AllContents);

    public XLColumn? FirstColumnUsed(XLCellsUsedOptions options)
    {
        XLRangeColumn? rngColumn = this.AsRange().FirstColumnUsed(options);
        return rngColumn != null
            ? this.Column(rngColumn.RangeAddress.FirstAddress.ColumnNumber)
            : null;
    }

    public XLColumn? LastColumnUsed() => this.LastColumnUsed(XLCellsUsedOptions.AllContents);

    public XLColumn? LastColumnUsed(XLCellsUsedOptions options)
    {
        XLRangeColumn? rngColumn = this.AsRange().LastColumnUsed(options);
        return rngColumn != null
            ? this.Column(rngColumn.RangeAddress.LastAddress.ColumnNumber)
            : null;
    }

    public XLRow Row(int row) => this.Row(row, true);

    public XLColumn Column(int columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnNumber),
                $"Column number must be between 1 and {XlsxSharp.XLHelper.MaxColumnNumber}"
            );
        }

        if (this.Internals.ColumnsCollection.TryGetValue(columnNumber, out XLColumn column))
        {
            return column;
        }
        else
        {
            // This is a new column so we're going to reference all
            // cells in this column to preserve their formatting
            this.Internals.RowsCollection.Keys.ForEach(r => this.Cell(r, columnNumber).PingStyle());

            column = this.RangeFactory.CreateColumn(columnNumber);
            this.Internals.ColumnsCollection.Add(columnNumber, column);
        }

        return column;
    }

    public IXLColumn Column(string column) =>
        this.Column(XlsxSharp.XLHelper.GetColumnNumberFromLetter(column));

    public override XLRange AsRange() =>
        this.Range(1, 1, XlsxSharp.XLHelper.MaxRowNumber, XlsxSharp.XLHelper.MaxColumnNumber);

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        if (!range.IsEntireColumn())
        {
            XLRangeAddress model = new(
                range.RangeAddress.FirstAddress,
                new XLAddress(
                    range.RangeAddress.LastAddress.RowNumber,
                    XlsxSharp.XLHelper.MaxColumnNumber,
                    false,
                    false
                )
            );
            List<IXLRange> rangesToSplit =
            [
                .. this
                    .Worksheet.MergedRanges.GetIntersectedRanges(model)
                    .Where(r =>
                        r.RangeAddress.FirstAddress.RowNumber
                            < range.RangeAddress.FirstAddress.RowNumber
                        || r.RangeAddress.LastAddress.RowNumber
                            > range.RangeAddress.LastAddress.RowNumber
                    ),
            ];
            foreach (IXLRange rangeToSplit in rangesToSplit)
            {
                this.Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        this.ShiftPageBreaksColumns(range, columnsShifted);

        List<ISheetListener> sheetListeners = this.GetSheetListeners();

        if (columnsShifted > 0)
        {
            Area insertedArea = Excel
                .Area.FromRangeAddress(range.RangeAddress)
                .SliceFromLeft(1)
                .ExtendRight(columnsShifted - 1);
            foreach (ISheetListener listener in sheetListeners)
            {
                listener.OnInsertAreaAndShiftRight(range.Worksheet, insertedArea);
            }
        }
        else if (columnsShifted < 0)
        {
            Area area = Excel.Area.FromRangeAddress(range.RangeAddress);
            foreach (ISheetListener listener in sheetListeners)
            {
                listener.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
            }
        }
    }

    private void ShiftPageBreaksColumns(XLRange range, int columnsShifted)
    {
        for (int i = 0; i < this.PageSetup.ColumnBreaks.Count; i++)
        {
            int br = this.PageSetup.ColumnBreaks[i];
            if (range.RangeAddress.FirstAddress.ColumnNumber <= br)
            {
                this.PageSetup.ColumnBreaks[i] = br + columnsShifted;
            }
        }
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        if (!range.IsEntireRow())
        {
            XLRangeAddress model = new(
                range.RangeAddress.FirstAddress,
                new XLAddress(
                    XlsxSharp.XLHelper.MaxRowNumber,
                    range.RangeAddress.LastAddress.ColumnNumber,
                    false,
                    false
                )
            );
            List<IXLRange> rangesToSplit =
            [
                .. this
                    .Worksheet.MergedRanges.GetIntersectedRanges(model)
                    .Where(r =>
                        r.RangeAddress.FirstAddress.ColumnNumber
                            < range.RangeAddress.FirstAddress.ColumnNumber
                        || r.RangeAddress.LastAddress.ColumnNumber
                            > range.RangeAddress.LastAddress.ColumnNumber
                    ),
            ];
            foreach (IXLRange rangeToSplit in rangesToSplit)
            {
                this.Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        this.ShiftPageBreaksRows(range, rowsShifted);

        List<ISheetListener> sheetListeners = this.GetSheetListeners();

        if (rowsShifted > 0)
        {
            Area insertedArea = Excel
                .Area.FromRangeAddress(range.RangeAddress)
                .SliceFromTop(1)
                .ExtendBelow(rowsShifted - 1);
            foreach (ISheetListener listener in sheetListeners)
            {
                listener.OnInsertAreaAndShiftDown(range.Worksheet, insertedArea);
            }
        }
        else if (rowsShifted < 0)
        {
            Area area = Excel.Area.FromRangeAddress(range.RangeAddress);
            foreach (ISheetListener listener in sheetListeners)
            {
                listener.OnDeleteAreaAndShiftUp(range.Worksheet, area);
            }
        }
    }

    private List<ISheetListener> GetSheetListeners()
    {
        List<ISheetListener> sheetListeners =
        [
            this.Workbook.CalcEngine,
            this.Hyperlinks,
            this.Workbook.DefinedNamesInternal,
            this.DataValidations,
        ];
        foreach (XLWorksheet worksheet in this.Workbook.WorksheetsInternal)
        {
            sheetListeners.Add(worksheet.DefinedNames);
        }

        sheetListeners.AddRange(this.SparklineGroupsInternal);

        // CF can contain formulas for any worksheet, notify about all changes
        foreach (XLWorksheet worksheet in this.Workbook.WorksheetsInternal)
        {
            sheetListeners.Add(worksheet.ConditionalFormats);
        }

        return sheetListeners;
    }

    private void ShiftPageBreaksRows(XLRange range, int rowsShifted)
    {
        for (int i = 0; i < this.PageSetup.RowBreaks.Count; i++)
        {
            int br = this.PageSetup.RowBreaks[i];
            if (range.RangeAddress.FirstAddress.RowNumber <= br)
            {
                this.PageSetup.RowBreaks[i] = br + rowsShifted;
            }
        }
    }

    public void NotifyRangeShiftedRows(XLRange range, int rowsShifted)
    {
        List<XLRangeBase> rangesToShift =
        [
            .. this
                ._rangeRepository.Where(r => r.RangeAddress.IsValid)
                .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber * -Math.Sign(rowsShifted)),
        ];

        this.WorksheetRangeShiftedRows(range, rowsShifted);

        bool collapsed = false;
        foreach (XLRangeBase storedRange in rangesToShift)
        {
            if (storedRange.IsEntireColumn())
            {
                continue;
            }

            if (ReferenceEquals(range, storedRange))
            {
                continue;
            }

            storedRange.WorksheetRangeShiftedRows(range, rowsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedRows(range, rowsShifted);
        }
    }

    public void NotifyRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        List<XLRangeBase> rangesToShift =
        [
            .. this
                ._rangeRepository.Where(r => r.RangeAddress.IsValid)
                .OrderBy(r =>
                    r.RangeAddress.FirstAddress.ColumnNumber * -Math.Sign(columnsShifted)
                ),
        ];

        this.WorksheetRangeShiftedColumns(range, columnsShifted);

        bool collapsed = false;
        foreach (XLRangeBase storedRange in rangesToShift)
        {
            if (storedRange.IsEntireRow())
            {
                continue;
            }

            if (ReferenceEquals(range, storedRange))
            {
                continue;
            }

            storedRange.WorksheetRangeShiftedColumns(range, columnsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedColumns(range, columnsShifted);
        }
    }

    internal XLRow Row(int rowNumber, bool pingCells)
    {
        if (rowNumber <= 0 || rowNumber > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(rowNumber),
                $"Row number must be between 1 and {XlsxSharp.XLHelper.MaxRowNumber}"
            );
        }

        if (this.Internals.RowsCollection.TryGetValue(rowNumber, out XLRow row))
        {
            return row;
        }
        else
        {
            if (pingCells)
            {
                // This is a new row so we're going to reference all
                // cells in columns of this row to preserve their formatting
                this.Internals.ColumnsCollection.Keys.ForEach(c =>
                    this.Cell(rowNumber, c).PingStyle()
                );
            }

            row = this.RangeFactory.CreateRow(rowNumber);
            this.Internals.RowsCollection.Add(rowNumber, row);
        }

        return row;
    }

    public IXLTable Table(XLRange range, bool addToTables, bool setAutofilter = true) =>
        this.Table(
            range,
            TableNameGenerator.GetNewTableName(this.Workbook),
            addToTables,
            setAutofilter
        );

    public IXLTable Table(XLRange range, string name, bool addToTables, bool setAutofilter = true)
    {
        this.CheckRangeNotOverlappingOtherEntities(range);
        XLRangeAddress rangeAddress;
        if (range.Rows().Count() == 1)
        {
            rangeAddress = new XLRangeAddress(
                range.FirstCell().Address,
                range.LastCell().CellBelow().Address
            );
            range.InsertRowsBelow(1);
        }
        else
        {
            rangeAddress = range.RangeAddress;
        }

        XLRangeKey rangeKey = new(XLRangeType.Table, rangeAddress);
        XLTable table = (XLTable)this._rangeRepository.GetOrCreate(ref rangeKey);

        if (table.Name != name)
        {
            table.Name = name;
        }

        if (addToTables && !this.Tables.Contains(table))
        {
            this.Tables.Add(table);
        }

        if (setAutofilter && !table.ShowAutoFilter)
        {
            table.InitializeAutoFilter();
        }

        return table;
    }

    private void CheckRangeNotOverlappingOtherEntities(XLRange range)
    {
        // Check that the range doesn't overlap with any existing tables
        XLTable? firstOverlappingTable = this.Tables.FirstOrDefault<XLTable>(t =>
            t.RangeUsed().Intersects(range)
        );
        if (firstOverlappingTable != null)
        {
            throw new InvalidOperationException(
                $"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'"
            );
        }

        // Check that the range doesn't overlap with any filters
        if (this.AutoFilter.IsEnabled && this.AutoFilter.Range.Intersects(range))
        {
            throw new InvalidOperationException(
                $"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} overlaps with the worksheet's autofilter."
            );
        }
    }

    private IXLRange GetRangeForSort()
    {
        IXLRange range = this.RangeUsed();
        this.SortColumns.ForEach(e =>
            range.SortColumns.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase)
        );
        this.SortRows.ForEach(e =>
            range.SortRows.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase)
        );
        return range;
    }

    public XLPivotTable PivotTable(string name) => (XLPivotTable)this.PivotTables.PivotTable(name);

    public override IXLCells Cells() => this.Cells(true, XLCellsUsedOptions.All);

    public override XLCells Cells(bool usedCellsOnly)
    {
        if (usedCellsOnly)
        {
            return this.Cells(true, XLCellsUsedOptions.AllContents);
        }
        else
        {
            return this.Range(
                    (this as IXLRangeBase).FirstCellUsed(XLCellsUsedOptions.All),
                    (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All)
                )
                .Cells(false, XLCellsUsedOptions.All);
        }
    }

    public override XLCell? Cell(string cellAddressInRange)
    {
        XLCell? cell = base.Cell(cellAddressInRange);
        if (cell is not null)
        {
            return cell;
        }

        if (
            this.Workbook.DefinedNames.TryGetValue(
                cellAddressInRange,
                out IXLDefinedName? definedName
            )
        )
        {
            if (!definedName.Ranges.Any())
            {
                return null;
            }

            return definedName.Ranges.First().FirstCell().CastTo<XLCell>();
        }

        return null;
    }

    public override XLRange? Range(string rangeAddressStr)
    {
        if (XlsxSharp.XLHelper.IsValidRangeAddress(rangeAddressStr))
        {
            return this.Range(new XLRangeAddress(this.Worksheet, rangeAddressStr));
        }

        if (rangeAddressStr.Contains('['))
        {
            return this.Table(rangeAddressStr.Substring(0, rangeAddressStr.IndexOf("[")))
                as XLRange;
        }

        if (this.DefinedNames.TryGetValue(rangeAddressStr, out IXLDefinedName? sheetDefinedName))
        {
            return sheetDefinedName.Ranges.First().CastTo<XLRange>();
        }

        if (
            this.Workbook.DefinedNamesInternal.TryGetValue(
                rangeAddressStr,
                out IXLDefinedName? workbookDefinedName
            )
        )
        {
            if (!workbookDefinedName.Ranges.Any())
            {
                return null;
            }

            return workbookDefinedName.Ranges.First().CastTo<XLRange>();
        }

        return null;
    }

    public IXLRanges MergedRanges => this.Internals.MergedRanges;

    IXLConditionalFormats IXLWorksheet.ConditionalFormats => this.ConditionalFormats;

    internal XLConditionalFormats ConditionalFormats { get; }

    public IXLSparklineGroups SparklineGroups => this.SparklineGroupsInternal;

    public IXLRanges SelectedRanges
    {
        get
        {
            this._selectedRanges.RemoveAll(r => !r.RangeAddress.IsValid);
            return this._selectedRanges;
        }
    }

    IXLCell? IXLWorksheet.ActiveCell
    {
        get => this.ActiveCell is not null ? new XLCell(this, this.ActiveCell.Value) : null;
        set => this.ActiveCell = value is not null ? Point.FromAddress(value.Address) : null;
    }

    /// <summary>
    /// Address of active cell/cursor in the worksheet.
    /// </summary>
    internal Point? ActiveCell { get; set; }

    private XLCalcEngine CalcEngine => this.Workbook.CalcEngine;

    public XLCellValue Evaluate(string expression, string? formulaAddress = null)
    {
        IXLAddress? address = formulaAddress is not null ? XLAddress.Create(formulaAddress) : null;
        return this
            .CalcEngine.EvaluateFormula(expression, this.Workbook, this, address, true)
            .ToCellValue();
    }

    public void RecalculateAllFormulas()
    {
        this.Internals.CellsCollection.FormulaSlice.MarkDirty(Excel.Area.Full);
        this.Workbook.CalcEngine.Recalculate(this.Workbook, this.Name);
    }

    public string Author { get; set; }

    public override string ToString() => this.Name;

    IXLPictures IXLWorksheet.Pictures => this.Pictures;

    public bool IsPasswordProtected => this.Protection.IsPasswordProtected;

    public bool IsProtected => this.Protection.IsProtected;

    internal XLPictures Pictures { get; }

    public IXLPicture Picture(string pictureName) => this.Pictures.Picture(pictureName);

    public IXLPicture AddPicture(Stream stream) => this.Pictures.Add(stream);

    public IXLPicture AddPicture(Stream stream, string name) => this.Pictures.Add(stream, name);

    internal IXLPicture AddPicture(Stream stream, string name, int Id) =>
        ((XLPictures)this.Pictures).Add(stream, name, Id);

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format) =>
        this.Pictures.Add(stream, format);

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name) =>
        this.Pictures.Add(stream, format, name);

    public IXLPicture AddPicture(string imageFile) => this.Pictures.Add(imageFile);

    public IXLPicture AddPicture(string imageFile, string name) =>
        this.Pictures.Add(imageFile, name);

    public override bool IsEntireRow() => true;

    public override bool IsEntireColumn() => true;

    internal IXLTable InsertTable(
        Point origin,
        IInsertDataReader reader,
        string tableName,
        bool createTable,
        bool addHeadings,
        bool transpose
    )
    {
        if (createTable && this.Tables.Any<XLTable>(t => t.Area.Contains(origin)))
        {
            throw new InvalidOperationException(
                $"This cell '{origin}' is already part of a table."
            );
        }

        XLRange range = this.InsertData(origin, reader, addHeadings, transpose);

        if (createTable)
        // Create a table and save it in the file
        {
            return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
        }
        else
        // Create a table, but keep it in memory. Saved file will contain only "raw" data and column headers
        {
            return tableName == null ? range.AsTable() : range.AsTable(tableName);
        }
    }

    internal XLRange InsertData(
        Point origin,
        IInsertDataReader reader,
        bool addHeadings,
        bool transpose
    )
    {
        // Prepare data. Heading is basically just another row of data, so unify it.
        IEnumerable<IEnumerable<XLCellValue>> rows = reader.GetRecords();
        int propCount = reader.GetPropertiesCount();
        if (addHeadings)
        {
            XLCellValue[] headings = new XLCellValue[propCount];
            for (int i = 0; i < propCount; i++)
            {
                headings[i] = reader.GetPropertyName(i);
            }

            rows = new[] { headings }.Concat(rows);
        }

        if (transpose)
        {
            rows = TransposeJaggedArray(rows);
        }

        ValueSlice valueSlice = this.Internals.CellsCollection.ValueSlice;
        FormatSlice formatSlice = this.Internals.CellsCollection.FormatSlice;

        // A buffer to avoid multiple enumerations of the source.
        List<XLCellValue> rowBuffer = [];
        int maximumColumn = origin.Column;
        int rowNumber = origin.Row;
        foreach (IEnumerable<XLCellValue> row in rows)
        {
            rowBuffer.AddRange(row);

            // InsertData should also clear data and if row doesn't have enough data,
            // fill in the rest. Only fill up to the props to be consistent. We can't
            // know how long any next row will be, so props are used as a source of truth
            // for which columns should be cleared.
            for (int i = rowBuffer.Count; i < propCount; ++i)
            {
                rowBuffer.Add(Blank.Value);
            }

            // Each row can have different number of values, so we have to check every row.
            maximumColumn = Math.Max(origin.Column + rowBuffer.Count - 1, maximumColumn);
            if (
                maximumColumn > XlsxSharp.XLHelper.MaxColumnNumber
                || rowNumber > XlsxSharp.XLHelper.MaxRowNumber
            )
            {
                throw new ArgumentException("Data would write out of the sheet.");
            }

            int column = origin.Column;
            for (int i = 0; i < rowBuffer.Count; ++i)
            {
                XLCellValue value = rowBuffer[i];
                Point point = new(rowNumber, column);
                XLCellFormatValue? modifiedStyle = this.GetStyleForValue(value, point);
                if (modifiedStyle is not null)
                {
                    if (value.IsText && value.GetText()[0] == '\'')
                    {
                        value = value.GetText().Substring(1);
                    }

                    formatSlice.Set(point, modifiedStyle);
                }

                valueSlice.SetCellValue(point, value);
                column++;
            }

            rowBuffer.Clear();
            rowNumber++;
        }

        // If there is no row, rowNumber is kept at origin instead of last row + 1 .
        int lastRow = Math.Max(rowNumber - 1, origin.Row);
        Area insertedArea = new(origin, new Point(lastRow, maximumColumn));

        // If inserted area affected a table, we must fix headings and totals, because these values
        // are duplicated. Basically the table values are the truth and cells are a reflection of the
        // truth, but here we inserted shadow first.
        foreach (XLTable table in this.Tables)
        {
            table.RefreshFieldsFromCells(insertedArea);
        }

        // Invalidate only once, not for every cell.
        this.Workbook.CalcEngine.MarkDirty(this.Worksheet, insertedArea);

        // Return area that contains all inserted cells, no matter how jagged were data.
        return this.Range(
            insertedArea.FirstPoint.Row,
            insertedArea.FirstPoint.Column,
            insertedArea.LastPoint.Row,
            insertedArea.LastPoint.Column
        );

        // Rather memory inefficient, but the original code also materialized
        // data through Linq/required multiple enumerations.
        static List<List<XLCellValue>> TransposeJaggedArray(
            IEnumerable<IEnumerable<XLCellValue>> enumerable
        )
        {
            List<List<XLCellValue>> destination = [];

            int sourceRow = 1;
            foreach (IEnumerable<XLCellValue> row in enumerable)
            {
                int sourceColumn = 1;
                foreach (XLCellValue sourceValue in row)
                {
                    // The original has `sourceValue` at [sourceRow, sourceColumn]
                    int destinationRowCount = destination.Count;
                    if (sourceColumn > destinationRowCount)
                    {
                        destination.Add([]);
                    }

                    // There can be jagged arrays and the destination can have spaces between columns.
                    List<XLCellValue> destinationRow = destination[sourceColumn - 1];
                    while (destinationRow.Count < sourceRow - 1)
                    {
                        destinationRow.Add(Blank.Value);
                    }

                    destinationRow.Add(sourceValue);
                    sourceColumn++;
                }

                sourceRow++;
            }

            return destination;
        }
    }

    /// <summary>
    /// Get cell or null, if cell doesn't exist.
    /// </summary>
    internal XLCell? GetCell(Point point) =>
        this.Worksheet.Internals.CellsCollection.GetUsedCell(point);

    public XLRange GetOrCreateRange(XLRangeAddress rangeAddress)
    {
        XLRangeKey rangeKey = new(XLRangeType.Range, rangeAddress);
        XLRangeBase range = this._rangeRepository.GetOrCreate(ref rangeKey);

        return (XLRange)range;
    }

    /// <summary>
    /// Get a range row from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range row.</param>
    /// <returns>Range row with the specified address.</returns>
    public XLRangeRow RangeRow(XLRangeAddress address)
    {
        XLRangeKey rangeKey = new(XLRangeType.RangeRow, address);
        XLRangeRow rangeRow = (XLRangeRow)this._rangeRepository.GetOrCreate(ref rangeKey);
        return rangeRow;
    }

    /// <summary>
    /// Get a range column from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range column.</param>
    /// <returns>Range column with the specified address.</returns>
    public XLRangeColumn RangeColumn(XLRangeAddress address)
    {
        XLRangeKey rangeKey = new(XLRangeType.RangeColumn, address);
        XLRangeColumn rangeColumn = (XLRangeColumn)this._rangeRepository.GetOrCreate(ref rangeKey);
        return rangeColumn;
    }

    protected override void OnRangeAddressChanged(
        XLRangeAddress oldAddress,
        XLRangeAddress newAddress
    ) { }

    public void RelocateRange(
        XLRangeType rangeType,
        XLRangeAddress oldAddress,
        XLRangeAddress newAddress
    )
    {
        XLRangeKey oldKey = new(rangeType, oldAddress);
        XLRangeKey newKey = new(rangeType, newAddress);
        XLRangeBase? range = this._rangeRepository.Replace(ref oldKey, ref newKey);

        foreach (IXLRangeIndex rangeIndex in this._rangeIndices)
        {
            if (!rangeIndex.MatchesType(rangeType))
            {
                continue;
            }

            if (rangeIndex.Remove(oldAddress) && newAddress.IsValid && range != null)
            {
                rangeIndex.Add(range);
            }
        }
    }

    internal void DeleteColumn(int columnNumber)
    {
        this.Internals.ColumnsCollection.Remove(columnNumber);

        List<int> columnsToMove =
        [
            .. this
                .Internals.ColumnsCollection.Where(c => c.Key > columnNumber)
                .Select(c => c.Key)
                .OrderBy(c => c),
        ];
        foreach (int column in columnsToMove)
        {
            this.Internals.ColumnsCollection.Add(
                column - 1,
                this.Internals.ColumnsCollection[column]
            );
            this.Internals.ColumnsCollection.Remove(column);

            this.Internals.ColumnsCollection[column - 1].SetColumnNumber(column - 1);
        }
    }

    internal void DeleteRow(int rowNumber)
    {
        this.Internals.RowsCollection.Remove(rowNumber);

        List<int> rowsToMove =
        [
            .. this
                .Internals.RowsCollection.Where(c => c.Key > rowNumber)
                .Select(c => c.Key)
                .OrderBy(r => r),
        ];
        foreach (int row in rowsToMove)
        {
            this.Internals.RowsCollection.Add(
                row - 1,
                this.Worksheet.Internals.RowsCollection[row]
            );
            this.Internals.RowsCollection.Remove(row);

            this.Internals.RowsCollection[row - 1].SetRowNumber(row - 1);
        }
    }

    internal void DeleteRange(XLRangeAddress rangeAddress)
    {
        XLRangeKey rangeKey = new(XLRangeType.Range, rangeAddress);
        this._rangeRepository.Remove(ref rangeKey);
    }

    /// <summary>
    /// Get the actual style for a point in the sheet.
    /// </summary>
    internal XLCellFormatValue GetStyleValue(Point point)
    {
        // TODO Styles: This is basically a duplication of Hierarchy.Resolve(). Investigate deduplication.
        XLCellFormatValue? cellFormat = this.Internals.CellsCollection.FormatSlice.GetFormat(point);
        if (cellFormat is not null)
        {
            return cellFormat;
        }

        // TODO Styles: Ensure all cross points are set at this time (=load+change). Taking from row/col should only be done if no cross is there.
        // If the slice doesn't contain any value, determine values by inheriting.
        // Cells that lie on an intersection of a XLColumn and a XLRow have their
        // style set when column/row is created to avoid problems with correct which
        // style has precedence. I.e. set column blue, set row red => cell is red.
        // Swap order the the cell is blue.
        return this.GetInheritedFormat(point);
    }

    internal XLCellFormatValue GetInheritedFormat(Point point)
    {
        if (
            this.Internals.RowsCollection.TryGetValue(point.Row, out XLRow? row)
            && row.FormatValue is not null
        )
        {
            return row.FormatValue;
        }

        if (
            this.Internals.ColumnsCollection.TryGetValue(point.Column, out XLColumn? column)
            && column.FormatValue is not null
        )
        {
            return column.FormatValue;
        }

        XLCellFormatValue? sheetFormat = this.FormatValue;
        if (sheetFormat is not null)
        {
            return sheetFormat;
        }

        return this.Workbook.Styles.DefaultCellFormat;
    }

    /// <summary>
    /// Get a style that should be used for a <paramref name="value"/>,
    /// if the value is set to the <paramref name="point"/>.
    /// </summary>
    internal XLCellFormatValue? GetStyleForValue(XLCellValue value, Point point)
    {
        // Because StyleValue property retrieves value from a slice,
        // access it only if necessary. This happens during ever cell
        // of modification and thus is performance critical.
        switch (value.Type)
        {
            case XLDataType.DateTime:
                {
                    bool onlyDatePart = value.GetUnifiedNumber() % 1 == 0;
                    XLCellFormatValue currentFormat = this.GetStyleValue(point);
                    if (currentFormat.NumberFormat.IsGeneralFormat())
                    {
                        XLPredefinedFormat.DateTime numberFormatId = onlyDatePart
                            ? XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes
                            : XLPredefinedFormat.DateTime.MonthDayYear4WithDashesHour24Minutes;
                        XLNumberFormat dateTimeNumberFormat = XLPredefinedFormat.FormatCodes[
                            (int)numberFormatId
                        ];
                        return this.Workbook.Styles.GetModifiedFormat(
                            currentFormat,
                            dateTimeNumberFormat
                        );
                    }
                }
                break;

            case XLDataType.TimeSpan:
                {
                    XLCellFormatValue currentFormat = this.GetStyleValue(point);
                    if (currentFormat.NumberFormat.IsGeneralFormat())
                    {
                        XLNumberFormat durationNumberFormat = XLPredefinedFormat.FormatCodes[
                            (int)XLPredefinedFormat.DateTime.Hour12MinutesSeconds
                        ];
                        return this.Workbook.Styles.GetModifiedFormat(
                            currentFormat,
                            durationNumberFormat
                        );
                    }
                }
                break;

            case XLDataType.Text:
            {
                string text = value.GetText();
                bool startsWithQuote = text.Length > 0 && text[0] == '\'';
                bool containsNewLine = text.Contains(Environment.NewLine, StringComparison.Ordinal);
                if (!startsWithQuote && !containsNewLine)
                {
                    break;
                }

                XLCellFormatValue currentFormat = this.GetStyleValue(point);
                if (startsWithQuote && !currentFormat.IncludeQuotePrefix)
                {
                    currentFormat = this.Workbook.Styles.GetRegisteredCellFormat(
                        currentFormat,
                        format => format with { IncludeQuotePrefix = true }
                    );
                }

                if (containsNewLine && !currentFormat.Alignment.WrapText)
                {
                    currentFormat = this.Workbook.Styles.GetModifiedFormat(
                        currentFormat,
                        alignment => alignment with { WrapText = true }
                    );
                }

                return currentFormat;
            }
        }

        return null;
    }
}

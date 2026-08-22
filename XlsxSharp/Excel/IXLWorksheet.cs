using System;
using System.IO;
using XlsxSharp.Excel.CalcEngine.Exceptions;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

public enum XLWorksheetVisibility
{
    Visible,
    Hidden,
    VeryHidden,
}

public interface IXLWorksheet
    : IXLRangeBase,
        IXLProtectable<IXLSheetProtection, XLSheetProtectionElements>
{
    /// <summary>
    /// Gets the workbook that contains this worksheet
    /// </summary>
    public XLWorkbook Workbook { get; }

    /// <summary>
    /// Gets or sets the default column width for this worksheet.
    /// </summary>
    public double ColumnWidth { get; set; }

    /// <summary>
    /// Gets or sets the default row height for this worksheet.
    /// </summary>
    public double RowHeight { get; set; }

    /// <summary>
    /// Gets or sets the name (caption) of this worksheet. The sheet rename also renames sheet
    /// in formulas and defined names.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the position of the sheet.
    /// <para>When setting the Position all other sheets' positions are shifted accordingly.</para>
    /// </summary>
    public int Position { get; set; }

    /// <summary>
    /// Gets an object to manipulate the sheet's print options.
    /// </summary>
    public IXLPageSetup PageSetup { get; }

    /// <summary>
    /// Gets an object to manipulate the Outline levels.
    /// </summary>
    public IXLOutline Outline { get; }

    /// <summary>
    /// All hyperlinks in the sheet.
    /// </summary>
    public IXLHyperlinks Hyperlinks { get; }

    /// <summary>
    /// Gets the first row of the worksheet.
    /// </summary>
    public IXLRow FirstRow();

    /// <summary>
    /// Gets the first non-empty row of the worksheet that contains a cell with a value.
    /// <para>Formatted empty cells do not count.</para>
    /// </summary>
    public IXLRow? FirstRowUsed();

    /// <summary>
    /// Gets the first non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLRow? FirstRowUsed(XLCellsUsedOptions options);

    /// <summary>
    /// Gets the last row of the worksheet.
    /// </summary>
    public IXLRow LastRow();

    /// <summary>
    /// Gets the last non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    public IXLRow? LastRowUsed();

    /// <summary>
    /// Gets the last non-empty row of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLRow? LastRowUsed(XLCellsUsedOptions options);

    /// <summary>
    /// Gets the first column of the worksheet.
    /// </summary>
    public IXLColumn FirstColumn();

    /// <summary>
    /// Gets the first non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    public IXLColumn? FirstColumnUsed();

    /// <summary>
    /// Gets the first non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLColumn? FirstColumnUsed(XLCellsUsedOptions options);

    /// <summary>
    /// Gets the last column of the worksheet.
    /// </summary>
    public IXLColumn LastColumn();

    /// <summary>
    /// Gets the last non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    public IXLColumn? LastColumnUsed();

    /// <summary>
    /// Gets the last non-empty column of the worksheet that contains a cell with a value.
    /// </summary>
    /// <param name="options">The options to determine whether a cell is used.</param>
    public IXLColumn? LastColumnUsed(XLCellsUsedOptions options);

    /// <summary>
    /// Gets a collection of all columns in this worksheet.
    /// </summary>
    public IXLColumns Columns();

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet, separated by commas.
    /// <para>e.g. Columns("G:H"), Columns("10:11,13:14"), Columns("P:Q,S:T"), Columns("V")</para>
    /// </summary>
    /// <param name="columns">The columns to return.</param>
    public IXLColumns Columns(string columns);

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet.
    /// </summary>
    /// <param name="firstColumn">The first column to return.</param>
    /// <param name="lastColumn">The last column to return.</param>
    public IXLColumns Columns(string firstColumn, string lastColumn);

    /// <summary>
    /// Gets a collection of the specified columns in this worksheet.
    /// </summary>
    /// <param name="firstColumn">The first column to return.</param>
    /// <param name="lastColumn">The last column to return.</param>
    public IXLColumns Columns(int firstColumn, int lastColumn);

    /// <summary>
    /// Gets a collection of all rows in this worksheet.
    /// </summary>
    public IXLRows Rows();

    /// <summary>
    /// Gets a collection of the specified rows in this worksheet, separated by commas.
    /// <para>e.g. Rows("4:5"), Rows("7:8,10:11"), Rows("13")</para>
    /// </summary>
    /// <param name="rows">The rows to return.</param>
    public IXLRows Rows(string rows);

    /// <summary>
    /// Gets a collection of the specified rows in this worksheet.
    /// </summary>
    /// <param name="firstRow">The first row to return.</param>
    /// <param name="lastRow">The last row to return.</param>
    public IXLRows Rows(int firstRow, int lastRow);

    /// <summary>
    /// Gets the specified row of the worksheet.
    /// </summary>
    /// <param name="row">The worksheet's row.</param>
    public IXLRow Row(int row);

    /// <summary>
    /// Gets the specified column of the worksheet.
    /// </summary>
    /// <param name="column">The worksheet's column.</param>
    public IXLColumn Column(int column);

    /// <summary>
    /// Gets the specified column of the worksheet.
    /// </summary>
    /// <param name="column">The worksheet's column.</param>
    public IXLColumn Column(string column);

    /// <summary>
    /// Gets the cell at the specified row and column.
    /// </summary>
    /// <param name="row">The cell's row.</param>
    /// <param name="column">The cell's column.</param>
    public IXLCell Cell(int row, int column);

    /// <summary>Gets the cell at the specified address.</summary>
    /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
    /// <exception cref="ArgumentException">Address is not A1 or workbook-scoped named range.</exception>
    public IXLCell Cell(string cellAddressInRange);

    /// <summary>
    /// Gets the cell at the specified row and column.
    /// </summary>
    /// <param name="row">The cell's row.</param>
    /// <param name="column">The cell's column.</param>
    public IXLCell Cell(int row, string column);

    /// <summary>Gets the cell at the specified address.</summary>
    /// <param name="cellAddressInRange">The cell address in the worksheet.</param>
    public IXLCell Cell(IXLAddress cellAddressInRange);

    /// <summary>
    /// Returns the specified range.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    public IXLRange Range(IXLRangeAddress rangeAddress);

    /// <summary>Returns the specified range.</summary>
    /// <para>e.g. Range("A1"), Range("A1:C2")</para>
    /// <param name="rangeAddress">The range boundaries.</param>
    /// <exception cref="ArgumentException"><paramref name="rangeAddress"/> is not a valid address or named range.</exception>
    public IXLRange Range(string rangeAddress);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCell">The first cell in the range.</param>
    /// <param name="lastCell"> The last cell in the range.</param>
    public IXLRange Range(IXLCell firstCell, IXLCell lastCell);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
    public IXLRange Range(string firstCellAddress, string lastCellAddress);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress"> The last cell address in the worksheet.</param>
    public IXLRange Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress);

    /// <summary>Returns a collection of ranges, separated by commas.</summary>
    /// <para>e.g. Ranges("A1"), Ranges("A1:C2"), Ranges("A1:B2,D1:D4")</para>
    /// <param name="ranges">The ranges to return.</param>
    public IXLRanges Ranges(string ranges);

    /// <summary>Returns the specified range.</summary>
    /// <param name="firstCellRow">   The first cell's row of the range to return.</param>
    /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
    /// <param name="lastCellRow">    The last cell's row of the range to return.</param>
    /// <param name="lastCellColumn"> The last cell's column of the range to return.</param>
    /// <returns>.</returns>
    public IXLRange Range(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    );

    /// <summary>Gets the number of rows in this worksheet.</summary>
    public int RowCount();

    /// <summary>Gets the number of columns in this worksheet.</summary>
    public int ColumnCount();

    /// <summary>
    /// Collapses all outlined rows.
    /// </summary>
    public IXLWorksheet CollapseRows();

    /// <summary>
    /// Collapses all outlined columns.
    /// </summary>
    public IXLWorksheet CollapseColumns();

    /// <summary>
    /// Expands all outlined rows.
    /// </summary>
    public IXLWorksheet ExpandRows();

    /// <summary>
    /// Expands all outlined columns.
    /// </summary>
    public IXLWorksheet ExpandColumns();

    /// <summary>
    /// Collapses the outlined rows of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLWorksheet CollapseRows(int outlineLevel);

    /// <summary>
    /// Collapses the outlined columns of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLWorksheet CollapseColumns(int outlineLevel);

    /// <summary>
    /// Expands the outlined rows of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLWorksheet ExpandRows(int outlineLevel);

    /// <summary>
    /// Expands the outlined columns of the specified level.
    /// </summary>
    /// <param name="outlineLevel">The outline level.</param>
    public IXLWorksheet ExpandColumns(int outlineLevel);

    /// <summary>
    /// Deletes this worksheet.
    /// </summary>
    public void Delete();

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    public IXLDefinedNames NamedRanges { get; }

    /// <summary>
    /// Gets an object to manage this worksheet's defined names.
    /// </summary>
    public IXLDefinedNames DefinedNames { get; }

    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXLDefinedName NamedRange(string rangeName);

    /// <summary>
    /// Gets the specified defined name.
    /// </summary>
    /// <param name="name">Name identifier of defined name, without sheet name.</param>
    /// <exception cref="ArgumentException">Name wasn't found in sheets defined names.</exception>
    public IXLDefinedName DefinedName(string name);

    /// <summary>
    /// Gets an object to manage how the worksheet is going to displayed by Excel.
    /// </summary>
    public IXLSheetView SheetView { get; }

    /// <summary>
    /// Gets the Excel table of the given index
    /// </summary>
    /// <param name="index">Index of the table to return</param>
    public IXLTable Table(int index);

    /// <summary>
    /// Gets the Excel table of the given name
    /// </summary>
    /// <param name="name">Name of the table to return</param>
    public IXLTable Table(string name);

    /// <summary>
    /// Gets an object to manage this worksheet's Excel tables
    /// </summary>
    public IXLTables Tables { get; }

    /// <summary>
    /// Copies the
    /// </summary>
    /// <param name="newSheetName"></param>
    public IXLWorksheet CopyTo(string newSheetName);

    public IXLWorksheet CopyTo(string newSheetName, int position);

    public IXLWorksheet CopyTo(XLWorkbook workbook);

    /// <summary>
    /// Copy a worksheet from this workbook to a different workbook as a new sheet.
    /// </summary>
    /// <param name="workbook">Workbook into which copy this sheet.</param>
    /// <param name="newSheetName">Name of new sheet in the <paramref name="workbook"/> where will the data be copied. Sheet will be in the last position.</param>
    /// <returns>Newly created sheet in the <paramref name="workbook"/>.</returns>
    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName);

    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName, int position);

    public IXLRange? RangeUsed();

    public IXLRange? RangeUsed(XLCellsUsedOptions options);

    public IXLDataValidations DataValidations { get; }

    public XLWorksheetVisibility Visibility { get; set; }

    public IXLWorksheet Hide();

    public IXLWorksheet Unhide();

    public IXLSortElements SortRows { get; }

    public IXLSortElements SortColumns { get; }

    public IXLRange Sort();

    public IXLRange Sort(
        string columnsToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRange Sort(
        int columnToSortBy,
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    public IXLRange SortLeftToRight(
        XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false,
        bool ignoreBlanks = true
    );

    //IXLCharts Charts { get; }

    public bool ShowFormulas { get; set; }

    public bool ShowGridLines { get; set; }

    public bool ShowOutlineSymbols { get; set; }

    public bool ShowRowColHeaders { get; set; }

    public bool ShowRuler { get; set; }

    public bool ShowWhiteSpace { get; set; }

    public bool ShowZeros { get; set; }

    public IXLWorksheet SetShowFormulas();
    public IXLWorksheet SetShowFormulas(bool value);

    public IXLWorksheet SetShowGridLines();
    public IXLWorksheet SetShowGridLines(bool value);

    public IXLWorksheet SetShowOutlineSymbols();
    public IXLWorksheet SetShowOutlineSymbols(bool value);

    public IXLWorksheet SetShowRowColHeaders();
    public IXLWorksheet SetShowRowColHeaders(bool value);

    public IXLWorksheet SetShowRuler();
    public IXLWorksheet SetShowRuler(bool value);

    public IXLWorksheet SetShowWhiteSpace();
    public IXLWorksheet SetShowWhiteSpace(bool value);

    public IXLWorksheet SetShowZeros();
    public IXLWorksheet SetShowZeros(bool value);

    public XLColor TabColor { get; set; }

    public IXLWorksheet SetTabColor(XLColor color);

    public bool TabSelected { get; set; }

    public bool TabActive { get; set; }

    public IXLWorksheet SetTabSelected();
    public IXLWorksheet SetTabSelected(bool value);

    public IXLWorksheet SetTabActive();
    public IXLWorksheet SetTabActive(bool value);

    public IXLPivotTable PivotTable(string name);

    public IXLPivotTables PivotTables { get; }

    public bool RightToLeft { get; set; }

    public IXLWorksheet SetRightToLeft();
    public IXLWorksheet SetRightToLeft(bool value);

    public IXLAutoFilter AutoFilter { get; }

    public IXLRows RowsUsed(
        XLCellsUsedOptions options = XLCellsUsedOptions.AllContents,
        Func<IXLRow, bool>? predicate = null
    );

    public IXLRows RowsUsed(Func<IXLRow, bool>? predicate);

    public IXLColumns ColumnsUsed(
        XLCellsUsedOptions options = XLCellsUsedOptions.AllContents,
        Func<IXLColumn, bool>? predicate = null
    );

    public IXLColumns ColumnsUsed(Func<IXLColumn, bool>? predicate);

    public IXLRanges MergedRanges { get; }

    public IXLConditionalFormats ConditionalFormats { get; }

    public IXLSparklineGroups SparklineGroups { get; }

    public IXLRanges SelectedRanges { get; }

    /// <summary>
    /// The active cell of the worksheet.
    /// </summary>
    public IXLCell? ActiveCell { get; set; }

    /// <summary>
    /// Evaluate an formula and return a result.
    /// </summary>
    /// <param name="expression">Formula to evaluate.</param>
    /// <param name="formulaAddress">A cell address that is used to provide context for formula calculation (mostly implicit intersection).</param>
    /// <exception cref="MissingContextException">If <paramref name="formulaAddress"/> was needed for some part of calculation.</exception>
    public XLCellValue Evaluate(string expression, string? formulaAddress = null);

    /// <summary>
    /// Force recalculation of all cell formulas in the sheet while leaving other sheets without change, even if their dirty cells.
    /// </summary>
    public void RecalculateAllFormulas();

    public string Author { get; set; }

    public IXLPictures Pictures { get; }

    public IXLPicture Picture(string pictureName);

    public IXLPicture AddPicture(Stream stream);

    public IXLPicture AddPicture(Stream stream, string name);

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format);

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name);

    public IXLPicture AddPicture(string imageFile);

    public IXLPicture AddPicture(string imageFile, string name);
}

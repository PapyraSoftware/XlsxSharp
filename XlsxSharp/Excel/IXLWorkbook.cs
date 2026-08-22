#nullable disable

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using XlsxSharp.Excel.CalcEngine.Exceptions;
using XlsxSharp.Excel.CustomProperties;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

public interface IXLWorkbook
    : IXLProtectable<IXLWorkbookProtection, XLWorkbookProtectionElements>,
        IDisposable
{
    public string Author { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's calculation mode.
    /// </summary>
    public XLCalculateMode CalculateMode { get; set; }

    public bool CalculationOnSave { get; set; }

    /// <summary>
    ///   Gets or sets the default column width for the workbook.
    ///   <para>All new worksheets will use this column width.</para>
    /// </summary>
    public double ColumnWidth { get; set; }

    public IXLCustomProperties CustomProperties { get; }

    public bool DefaultRightToLeft { get; }

    public bool DefaultShowFormulas { get; }

    public bool DefaultShowGridLines { get; }

    public bool DefaultShowOutlineSymbols { get; }

    public bool DefaultShowRowColHeaders { get; }

    public bool DefaultShowRuler { get; }

    public bool DefaultShowWhiteSpace { get; }

    public bool DefaultShowZeros { get; }

    public IXLFileSharing FileSharing { get; }

    public bool ForceFullCalculation { get; set; }

    public bool FullCalculationOnLoad { get; set; }

    public bool FullPrecision { get; set; }

    public bool LockStructure { get; set; }

    public bool LockWindows { get; set; }

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    public IXLDefinedNames NamedRanges { get; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's defined names.
    /// </summary>
    public IXLDefinedNames DefinedNames { get; }

    /// <summary>
    ///   Gets or sets the default outline options for the workbook.
    ///   <para>All new worksheets will use these outline options.</para>
    /// </summary>
    public IXLOutline Outline { get; set; }

    /// <summary>
    ///   Gets or sets the default page options for the workbook.
    ///   <para>All new worksheets will use these page options.</para>
    /// </summary>
    public IXLPageSetup PageOptions { get; set; }

    /// <summary>
    ///   Gets all pivot caches in a workbook. A one cache can be
    ///   used by multiple tables. Unused caches are not saved.
    /// </summary>
    public IXLPivotCaches PivotCaches { get; }

    /// <summary>
    ///   Gets or sets the workbook's properties.
    /// </summary>
    public XLWorkbookProperties Properties { get; set; }

    /// <summary>
    ///   Gets or sets the workbook's reference style.
    /// </summary>
    public XLReferenceStyle ReferenceStyle { get; set; }

    public bool RightToLeft { get; set; }

    /// <summary>
    ///   Gets or sets the default row height for the workbook.
    ///   <para>All new worksheets will use this row height.</para>
    /// </summary>
    public double RowHeight { get; set; }

    public bool ShowFormulas { get; set; }

    public bool ShowGridLines { get; set; }

    public bool ShowOutlineSymbols { get; set; }

    public bool ShowRowColHeaders { get; set; }

    public bool ShowRuler { get; set; }

    public bool ShowWhiteSpace { get; set; }

    public bool ShowZeros { get; set; }

    /// <summary>
    /// Gets or sets the default format of the workbook. All cells that don't have specified format, either
    /// at sheet, column, row or cell level, will use this format.
    /// </summary>
    public IXLStyle Style { get; set; }

    /// <summary>
    ///   Gets an object to manipulate this workbook's theme.
    /// </summary>
    public IXLTheme Theme { get; }

    public bool Use1904DateSystem { get; set; }

    /// <summary>
    ///   Gets an object to manipulate the worksheets.
    /// </summary>
    public IXLWorksheets Worksheets { get; }

    public IXLWorksheet AddWorksheet();

    public IXLWorksheet AddWorksheet(int position);

    public IXLWorksheet AddWorksheet(string sheetName);

    public IXLWorksheet AddWorksheet(string sheetName, int position);

    public void AddWorksheet(DataSet dataSet);

    public void AddWorksheet(IXLWorksheet worksheet);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The dataTable's name is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">Datatable to insert</param>
    /// <returns>Inserted Worksheet</returns>
    public IXLWorksheet AddWorksheet(DataTable dataTable);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1). The sheetName provided is used for the
    /// worksheet name. The name of a table will be generated as <em>Table{number suffix}</em>.
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet and Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName);

    /// <summary>
    /// Add a worksheet with a table at Cell(row:1, column:1).
    /// </summary>
    /// <param name="dataTable">dataTable to insert as Excel Table</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="tableName">Excel Table name</param>
    /// <returns>Inserted Worksheet</returns>
    public IXLWorksheet AddWorksheet(DataTable dataTable, string sheetName, string tableName);

    public IXLCell Cell(string namedCell);

    public IXLCells Cells(string namedCells);

    public IXLCustomProperty CustomProperty(string name);

    /// <summary>
    /// Evaluate a formula expression.
    /// </summary>
    /// <param name="expression">Formula expression to evaluate.</param>
    /// <exception cref="MissingContextException">
    /// If the expression contains a function that requires a context (e.g. current cell or worksheet).
    /// </exception>
    public XLCellValue Evaluate(string expression);

    public IXLCells FindCells(Func<IXLCell, bool> predicate);

    public IXLColumns FindColumns(Func<IXLColumn, bool> predicate);

    public IXLRows FindRows(Func<IXLRow, bool> predicate);

#nullable enable
    [Obsolete($"Use {nameof(DefinedName)} instead.")]
    public IXLDefinedName? NamedRange(string name);

    /// <summary>
    /// Try to find a defined name. If <paramref name="name"/> specifies a sheet, try to find
    /// name in the sheet first and fall back to the workbook if not found in the sheet.
    /// <para>
    /// <example>
    /// Requested name <c>Sheet1!Name</c> will first try to find <c>Name</c> in a sheet
    /// <c>Sheet1</c> (if such sheet exists) and if not found there, tries to find <c>Name</c>
    /// in workbook.
    /// </example>
    /// </para>
    /// <para>
    /// <example>
    /// Requested name <c>Name</c> will be searched only in a workbooks <see cref="DefinedNames"/>.
    /// </example>
    /// </para>
    /// </summary>
    /// <param name="name">Name of requested name, either plain name (e.g. <c>Name</c>) or with
    /// sheet specified (e.g. <c>Sheet!Name</c>).</param>
    /// <returns>Found name or null.</returns>
    public IXLDefinedName? DefinedName(string name);

#nullable disable

    public IXLRange Range(string range);

    public IXLRange RangeFromFullAddress(string rangeAddress, out IXLWorksheet ws);

    public IXLRanges Ranges(string ranges);

    /// <summary>
    /// Force recalculation of all cell formulas.
    /// </summary>
    public void RecalculateAllFormulas();

    /// <summary>
    ///   Saves the current workbook.
    /// </summary>
    public void Save();

    /// <summary>
    ///   Saves the current workbook and optionally performs validation
    /// </summary>
    public void Save(bool validate, bool evaluateFormulae = false);

    public void Save(SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a file.
    /// </summary>
    public void SaveAs(string file);

    /// <summary>
    ///   Saves the current workbook to a file and optionally validates it.
    /// </summary>
    public void SaveAs(string file, bool validate, bool evaluateFormulae = false);

    public void SaveAs(string file, SaveOptions options);

    /// <summary>
    ///   Saves the current workbook to a stream.
    /// </summary>
    public void SaveAs(Stream stream);

    /// <summary>
    ///   Saves the current workbook to a stream and optionally validates it.
    /// </summary>
    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false);

    public void SaveAs(Stream stream, SaveOptions options);

    /// <summary>
    /// Searches the cells' contents for a given piece of text
    /// </summary>
    /// <param name="searchText">The search text.</param>
    /// <param name="compareOptions">The compare options.</param>
    /// <param name="searchFormulae">if set to <c>true</c> search formulae instead of cell values.</param>
    public IEnumerable<IXLCell> Search(
        string searchText,
        CompareOptions compareOptions = CompareOptions.Ordinal,
        bool searchFormulae = false
    );

    public XLWorkbook SetLockStructure(bool value);

    public XLWorkbook SetLockWindows(bool value);

    public XLWorkbook SetUse1904DateSystem();

    public XLWorkbook SetUse1904DateSystem(bool value);

    /// <summary>
    /// Gets the Excel table of the given name
    /// </summary>
    /// <param name="tableName">Name of the table to return.</param>
    /// <param name="comparisonType">One of the enumeration values that specifies how the strings will be compared.</param>
    /// <returns>The table with given name</returns>
    /// <exception cref="ArgumentOutOfRangeException">If no tables with this name could be found in the workbook.</exception>
    public IXLTable Table(
        string tableName,
        StringComparison comparisonType = StringComparison.OrdinalIgnoreCase
    );

    public bool TryGetWorksheet(string name, out IXLWorksheet worksheet);

    public IXLWorksheet Worksheet(string name);

    public IXLWorksheet Worksheet(int position);
}

#nullable disable

using System.Collections;
using System.Data;

namespace XlsxSharp.Excel.Tables;

public interface IXLTable : IXLRange
{
    public IXLAutoFilter AutoFilter { get; }
    public IXLTableRange DataRange { get; }
    public bool EmphasizeFirstColumn { get; set; }
    public bool EmphasizeLastColumn { get; set; }
    public IEnumerable<IXLTableField> Fields { get; }

    /// <summary>
    /// Change the name of a table. Structural references to the table are not updated.
    /// </summary>
    /// <exception cref="ArgumentException">If the new table name is already used by other table in the sheet.</exception>
    public string Name { get; set; }

    public bool ShowAutoFilter { get; set; }
    public bool ShowColumnStripes { get; set; }
    public bool ShowHeaderRow { get; set; }
    public bool ShowRowStripes { get; set; }
    public bool ShowTotalsRow { get; set; }
    public XLTableTheme Theme { get; set; }

    /// <summary>
    /// Clears the contents of this table.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.All);

    /// <summary>
    /// Get field of the table.
    /// </summary>
    /// <param name="fieldName">Name of the field. Field names are case-insensitive.</param>
    /// <returns>Requested field.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Table doesn't contain <paramref name="fieldName"/> field.</exception>
    public IXLTableField Field(string fieldName);

    public IXLTableField Field(int fieldIndex);

    public IXLRangeRow HeadersRow();

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    public IXLRange AppendData(IEnumerable data, bool propagateExtraColumns = false);

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    public IXLRange AppendData(
        IEnumerable data,
        bool transpose,
        bool propagateExtraColumns = false
    );

    /// <summary>
    /// Appends the data of a data table and returns the range of the new rows.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    public IXLRange AppendData(DataTable dataTable, bool propagateExtraColumns = false);

    /// <summary>
    /// Appends the IEnumerable data elements and returns the range of the new rows.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="data">The table data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The range of the new rows.
    /// </returns>
    public IXLRange AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements and returns the table's data range.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    public IXLRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements and returns the table's data range.
    /// </summary>
    /// <param name="data">The IEnumerable data.</param>
    /// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    public IXLRange ReplaceData(
        IEnumerable data,
        bool transpose,
        bool propagateExtraColumns = false
    );

    /// <summary>
    /// Replaces the data from the records of a data table and returns the table's data range.
    /// </summary>
    /// <param name="dataTable">The data table.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    public IXLRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false);

    /// <summary>
    /// Replaces the IEnumerable data elements as a table and the table's data range.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="data">The table data.</param>
    /// <param name="propagateExtraColumns">if set to <c>true</c> propagate extra columns' values and formulas.</param>
    /// <returns>
    /// The table's data range.
    /// </returns>
    public IXLRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="range">The new table range.</param>
    public IXLTable Resize(IXLRange range);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    public IXLTable Resize(IXLRangeAddress rangeAddress);

    /// <summary>
    /// Resizes the table to the specified range address.
    /// </summary>
    /// <param name="rangeAddress">The range boundaries.</param>
    public IXLTable Resize(string rangeAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCell">The first cell in the range.</param>
    /// <param name="lastCell">The last cell in the range.</param>
    public IXLTable Resize(IXLCell firstCell, IXLCell lastCell);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
    public IXLTable Resize(string firstCellAddress, string lastCellAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellAddress">The first cell address in the worksheet.</param>
    /// <param name="lastCellAddress">The last cell address in the worksheet.</param>
    public IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress);

    /// <summary>
    /// Resizes the table to the specified range.
    /// </summary>
    /// <param name="firstCellRow">The first cell's row of the range to return.</param>
    /// <param name="firstCellColumn">The first cell's column of the range to return.</param>
    /// <param name="lastCellRow">The last cell's row of the range to return.</param>
    /// <param name="lastCellColumn">The last cell's column of the range to return.</param>
    public IXLTable Resize(
        int firstCellRow,
        int firstCellColumn,
        int lastCellRow,
        int lastCellColumn
    );

    public new IXLAutoFilter SetAutoFilter();

    public IXLTable SetEmphasizeFirstColumn();

    public IXLTable SetEmphasizeFirstColumn(bool value);

    public IXLTable SetEmphasizeLastColumn();

    public IXLTable SetEmphasizeLastColumn(bool value);

    public IXLTable SetShowAutoFilter();

    public IXLTable SetShowAutoFilter(bool value);

    public IXLTable SetShowColumnStripes();

    public IXLTable SetShowColumnStripes(bool value);

    public IXLTable SetShowHeaderRow();

    public IXLTable SetShowHeaderRow(bool value);

    public IXLTable SetShowRowStripes();

    public IXLTable SetShowRowStripes(bool value);

    public IXLTable SetShowTotalsRow();

    public IXLTable SetShowTotalsRow(bool value);

    public IXLRangeRow TotalsRow();

    /// <summary>
    /// Converts the table to an enumerable of dynamic objects
    /// </summary>
    public IEnumerable<dynamic> AsDynamicEnumerable();

    /// <summary>
    /// Converts the table to a standard .NET System.Data.DataTable
    /// </summary>
    public DataTable AsNativeDataTable();

    public IXLTable CopyTo(IXLWorksheet targetSheet);
}

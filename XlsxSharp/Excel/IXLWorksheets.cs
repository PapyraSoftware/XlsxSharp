using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

public interface IXLWorksheets : IEnumerable<IXLWorksheet>
{
    public int Count { get; }

    public IXLWorksheet Add();

    public IXLWorksheet Add(int position);

    public IXLWorksheet Add(string sheetName);

    public IXLWorksheet Add(string sheetName, int position);

    public IXLWorksheet Add(DataTable dataTable);

    public IXLWorksheet Add(DataTable dataTable, string sheetName);

    public IXLWorksheet Add(DataTable dataTable, string sheetName, string tableName);

    public void Add(DataSet dataSet);

    public bool Contains(string sheetName);

    public void Delete(string sheetName);

    public void Delete(int position);

    /// <summary>
    /// Try to get a sheet of a workbook with the specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <param name="worksheet">Found sheet or null if sheet is not found.</param>
    /// <returns><c>true</c> when sheet was found or <c>false</c> when it wasn't.</returns>
    public bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXLWorksheet? worksheet);

    /// <summary>
    /// Get a sheet of a workbook with specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <returns>Sheet with the specified name.</returns>
    /// <exception cref="KeyNotFoundException">When sheet with <paramref name="sheetName"/> isn't among the sheets.</exception>
    public IXLWorksheet Worksheet(string sheetName);

    public IXLWorksheet Worksheet(int position);
}

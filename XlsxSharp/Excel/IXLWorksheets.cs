using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

public interface IXLWorksheets : IEnumerable<IXLWorksheet>
{
    int Count { get; }

    IXLWorksheet Add();

    IXLWorksheet Add(int position);

    IXLWorksheet Add(string sheetName);

    IXLWorksheet Add(string sheetName, int position);

    IXLWorksheet Add(DataTable dataTable);

    IXLWorksheet Add(DataTable dataTable, string sheetName);

    IXLWorksheet Add(DataTable dataTable, string sheetName, string tableName);

    void Add(DataSet dataSet);

    bool Contains(string sheetName);

    void Delete(string sheetName);

    void Delete(int position);

    /// <summary>
    /// Try to get a sheet of a workbook with the specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <param name="worksheet">Found sheet or null if sheet is not found.</param>
    /// <returns><c>true</c> when sheet was found or <c>false</c> when it wasn't.</returns>
    bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXLWorksheet? worksheet);

    /// <summary>
    /// Get a sheet of a workbook with specified name. Sheet names are case-insensitive.
    /// </summary>
    /// <param name="sheetName">Name of sought sheet.</param>
    /// <returns>Sheet with the specified name.</returns>
    /// <exception cref="KeyNotFoundException">When sheet with <paramref name="sheetName"/> isn't among the sheets.</exception>
    IXLWorksheet Worksheet(string sheetName);

    IXLWorksheet Worksheet(int position);
}

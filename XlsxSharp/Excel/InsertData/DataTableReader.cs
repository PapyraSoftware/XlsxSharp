using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace XlsxSharp.Excel.InsertData;

internal class DataTableReader : IInsertDataReader
{
    private readonly IEnumerable<DataRow> _dataRows;
    private readonly DataTable? _dataTable;

    public DataTableReader(DataTable dataTable)
    {
        this._dataTable = dataTable ?? throw new ArgumentNullException(nameof(dataTable));
        this._dataRows = this._dataTable.Rows.Cast<DataRow>();
    }

    public DataTableReader(IEnumerable<DataRow> dataRows)
    {
        this._dataRows = dataRows ?? throw new ArgumentNullException(nameof(dataRows));
        this._dataTable = this._dataRows.FirstOrDefault()?.Table;
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords() =>
        this._dataRows.Select(r => r.ItemArray.Select(XLCellValue.FromInsertedObject));

    public int GetPropertiesCount()
    {
        if (this._dataTable != null)
        {
            return this._dataTable.Columns.Count;
        }

        if (this._dataRows.Any())
        {
            return this._dataRows.First().ItemArray.Length;
        }

        return 0;
    }

    public string? GetPropertyName(int propertyIndex)
    {
        if (propertyIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(propertyIndex),
                "Property index must be non-negative"
            );
        }

        if (this._dataTable == null)
        {
            return null;
        }

        if (propertyIndex >= this._dataTable.Columns.Count)
        {
            throw new ArgumentOutOfRangeException(
                $"{propertyIndex} exceeds the number of the table columns"
            );
        }

        return this._dataTable.Columns[propertyIndex].Caption;
    }
}

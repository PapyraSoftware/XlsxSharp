#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace XlsxSharp.Excel.InsertData;

internal class DataRecordReader : IInsertDataReader
{
    private readonly IEnumerable<XLCellValue>[] _inMemoryData;
    private string[] _columns;

    public DataRecordReader(IEnumerable<IDataRecord> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        this._inMemoryData = [.. this.ReadToEnd(data)];
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
    {
        return this._inMemoryData;
    }

    public int GetPropertiesCount()
    {
        return this._columns.Length;
    }

    public string GetPropertyName(int propertyIndex)
    {
        if (propertyIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(propertyIndex),
                "Property index must be non-negative"
            );
        }

        if (this._columns == null)
        {
            return null;
        }

        if (propertyIndex >= this._columns.Length)
        {
            throw new ArgumentOutOfRangeException(
                $"{propertyIndex} exceeds the number of the table columns"
            );
        }

        return this._columns[propertyIndex];
    }

    private IEnumerable<IEnumerable<XLCellValue>> ReadToEnd(IEnumerable<IDataRecord> data)
    {
        foreach (IDataRecord dataRecord in data)
        {
            yield return this.ToEnumerable(dataRecord).ToArray();
        }
    }

    private IEnumerable<XLCellValue> ToEnumerable(IDataRecord dataRecord)
    {
        bool firstRow = false;
        if (this._columns == null)
        {
            firstRow = true;
            this._columns = new string[dataRecord.FieldCount];
        }

        for (int i = 0; i < dataRecord.FieldCount; i++)
        {
            if (firstRow)
            {
                this._columns[i] = dataRecord.GetName(i);
            }

            object value = dataRecord[i];
            yield return XLCellValue.FromInsertedObject(value);
        }
    }
}

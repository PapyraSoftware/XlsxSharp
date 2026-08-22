using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.InsertData;

internal class ArrayReader : IInsertDataReader
{
    private readonly IEnumerable<IEnumerable> _data;

    public ArrayReader(IEnumerable<IEnumerable> data) =>
        this._data = data ?? throw new ArgumentNullException(nameof(data));

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords() =>
        this._data.Select(item => item.Cast<object>().Select(XLCellValue.FromInsertedObject));

    public int GetPropertiesCount()
    {
        if (!this._data.Any())
        {
            return 0;
        }

        return this._data.First().Cast<object>().Count();
    }

    public string? GetPropertyName(int propertyIndex) => null;
}

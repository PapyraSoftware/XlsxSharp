// Keep this file CodeMaid organised and cleaned

using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Excel.InsertData;

internal class NullDataReader : IInsertDataReader
{
    private readonly XLCellValue[] _row = [Blank.Value];
    private readonly int _count;

    public NullDataReader(IEnumerable<object> nulls) => this._count = nulls.Count();

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords() =>
        Enumerable.Repeat(this._row, this._count);

    public int GetPropertiesCount() => 0;

    public string? GetPropertyName(int propertyIndex) => null;
}

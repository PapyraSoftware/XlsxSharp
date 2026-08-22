#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

internal class XLPivotTables : IXLPivotTables, IEnumerable<XLPivotTable>
{
    private readonly Dictionary<String, XLPivotTable> _pivotTables = new(
        StringComparer.OrdinalIgnoreCase
    );

    public XLPivotTables(XLWorksheet worksheet)
    {
        this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    internal XLWorksheet Worksheet { get; }

    public void Add(XLPivotTable pivotTable)
    {
        XLPivotCache pivotCache = pivotTable.PivotCache;
        if (!pivotCache.FieldNames.Any())
        {
            pivotCache.Refresh();
        }

        this._pivotTables.Add(pivotTable.Name, pivotTable);
    }

    public IXLPivotTable Add(string name, IXLCell targetCell, IXLPivotCache pivotCache)
    {
        XLPivotTable pivotTable = new(this.Worksheet, (XLPivotCache)pivotCache)
        {
            Name = name,
            Area = new Area(Point.FromAddress(targetCell.Address)),
        };
        this.Add(pivotTable);
        pivotTable.UpdateCacheFields(Array.Empty<string>());
        return pivotTable;
    }

    public IXLPivotTable Add(string name, IXLCell targetCell, IXLRange range)
    {
        SheetArea area = SheetArea.From(range);
        XLPivotCaches pivotCaches = this.Worksheet.Workbook.PivotCachesInternal;
        XLPivotCache existingPivotCache = pivotCaches.Find(area);
        XLPivotCache pivotCache = existingPivotCache ?? pivotCaches.Add(area);
        return this.Add(name, targetCell, pivotCache);
    }

    public IXLPivotTable Add(string name, IXLCell targetCell, IXLTable table)
    {
        return this.Add(name, targetCell, (IXLRange)table);
    }

    public Boolean Contains(String name)
    {
        return this._pivotTables.ContainsKey(name);
    }

    public void Delete(String name)
    {
        this._pivotTables.Remove(name);
    }

    public void DeleteAll()
    {
        this._pivotTables.Clear();
    }

    IXLPivotTable IXLPivotTables.PivotTable(String name)
    {
        return this.PivotTable(name);
    }

    IEnumerator<IXLPivotTable> IEnumerable<IXLPivotTable>.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    IEnumerator<XLPivotTable> IEnumerable<XLPivotTable>.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public Dictionary<string, XLPivotTable>.ValueCollection.Enumerator GetEnumerator()
    {
        return this._pivotTables.Values.GetEnumerator();
    }

    internal void Add(String name, IXLPivotTable pivotTable)
    {
        this._pivotTables.Add(name, (XLPivotTable)pivotTable);
    }

    /// <inheritdoc cref="IXLPivotTables.PivotTable"/>
    internal XLPivotTable PivotTable(String name)
    {
        return this._pivotTables[name];
    }
}

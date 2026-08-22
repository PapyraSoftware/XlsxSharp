using System.Collections;
using System.Collections.Generic;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

internal class XLPivotCaches : IXLPivotCaches, IEnumerable<XLPivotCache>
{
    private readonly XLWorkbook _workbook;
    private readonly List<XLPivotCache> _caches = [];

    public XLPivotCaches(XLWorkbook workbook) => this._workbook = workbook;

    IXLPivotCache IXLPivotCaches.Add(IXLRange range) => this.Add(SheetArea.From(range));

    IEnumerator<IXLPivotCache> IEnumerable<IXLPivotCache>.GetEnumerator() => this.GetEnumerator();

    IEnumerator<XLPivotCache> IEnumerable<XLPivotCache>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public List<XLPivotCache>.Enumerator GetEnumerator() => this._caches.GetEnumerator();

    internal XLPivotCache Add(SheetArea area)
    {
        XLPivotSourceReference source = this._workbook.TryGetTable(area, out XLTable table)
            ? new XLPivotSourceReference(table.Name)
            : new XLPivotSourceReference(area);

        XLPivotCache newPivotCache = new(source, this._workbook);
        newPivotCache.Refresh();
        this._caches.Add(newPivotCache);
        return newPivotCache;
    }

    internal XLPivotCache Add(IXLPivotSource source)
    {
        XLPivotCache newPivotCache = new(source, this._workbook);
        this._caches.Add(newPivotCache);
        return newPivotCache;
    }

    /// <summary>
    /// Try to find an existing pivot cache for the passed area. The area
    /// is checked against both types of source references (tables and
    /// ranges) and if area matches, the cache is returned.
    /// </summary>
    internal XLPivotCache? Find(SheetArea area)
    {
        // This method mimics behavior of Excel.
        // If there is a table for the area and there is a cache for the table, return cache for the table.
        if (this._workbook.TryGetTable(area, out XLTable table))
        {
            // Table exists, so try to find it and match with the source reference.
            XLPivotSourceReference tableSource = new(table.Name);
            foreach (XLPivotCache cache in this._caches)
            {
                if (cache.Source.Equals(tableSource))
                {
                    return cache;
                }
            }
        }

        // Try to find a cache with area source.
        XLPivotSourceReference areaSource = new(area);
        foreach (XLPivotCache cache in this._caches)
        {
            if (cache.Source.Equals(areaSource))
            {
                return cache;
            }
        }

        return null;
    }
}

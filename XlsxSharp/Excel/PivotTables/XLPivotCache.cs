using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XlsxSharp.Excel.Exceptions;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLPivotCache : IXLPivotCache
{
    private readonly XLWorkbook _workbook;
    private readonly Dictionary<String, Int32> _fieldIndexes = new(XlsxSharp.XLHelper.NameComparer);
    private readonly List<String> _fieldNames = [];

    /// <summary>
    /// Length is a number of fields, in same order as <see cref="_fieldNames"/>.
    /// </summary>
    private readonly List<XLPivotCacheValues> _values = [];

    internal XLPivotCache(IXLPivotSource source, XLWorkbook workbook)
    {
        this._workbook = workbook;
        this.Guid = Guid.NewGuid();
        this.SetExcelDefaults();
        this.Source = source;
    }

    #region IXLPivotCache members

    public IReadOnlyList<String> FieldNames => this._fieldNames;

    public XLItemsToRetain ItemsToRetainPerField { get; set; }

    public Boolean RefreshDataOnOpen { get; set; }

    public Boolean SaveSourceData { get; set; }

    /// <summary>
    /// Number of fields in the cache.
    /// </summary>
    internal int FieldCount => this._fieldNames.Count;

    internal int RecordCount => this._fieldNames.Count > 0 ? this._values[0].Count : 0;

    public IXLPivotCache Refresh()
    {
        // Refresh can only happen if the reference is valid.
        if (!this.Source.TryGetSource(this._workbook, out XLWorksheet? sheet, out Area? foundArea))
        {
            throw new InvalidReferenceException();
        }

        Debug.Assert(sheet is not null && foundArea is not null);
        List<string> oldFieldNames = [.. this._fieldNames];
        this._fieldIndexes.Clear();
        this._fieldNames.Clear();
        this._values.Clear();

        ValueSlice valueSlice = sheet.Internals.CellsCollection.ValueSlice;
        Area area = foundArea.Value;
        for (int column = area.LeftColumn; column <= area.RightColumn; ++column)
        {
            string header = sheet.Cell(area.TopRow, column).GetFormattedString();

            XLPivotCacheValues fieldRecords = new(valueSlice, column, area);

            this.AddField(this.AdjustedFieldName(header), fieldRecords);
        }

        UpdatePivotTables();
        return this;

        void UpdatePivotTables()
        {
            foreach (XLWorksheet worksheet in this._workbook.WorksheetsInternal)
            {
                foreach (XLPivotTable pivotTable in worksheet.PivotTables)
                {
                    if (pivotTable.PivotCache == this)
                    {
                        pivotTable.UpdateCacheFields(oldFieldNames);
                    }
                }
            }
        }
    }

    public IXLPivotCache SetItemsToRetainPerField(XLItemsToRetain value)
    {
        this.ItemsToRetainPerField = value;
        return this;
    }

    public IXLPivotCache SetRefreshDataOnOpen() => this.SetRefreshDataOnOpen(true);

    public IXLPivotCache SetRefreshDataOnOpen(Boolean value)
    {
        this.RefreshDataOnOpen = value;
        return this;
    }

    public IXLPivotCache SetSaveSourceData() => this.SetSaveSourceData(true);

    public IXLPivotCache SetSaveSourceData(Boolean value)
    {
        this.SaveSourceData = value;
        return this;
    }

    #endregion

    /// <summary>
    /// Pivot cache definition id from the file.
    /// </summary>
    internal uint? CacheId { get; set; }

    internal Guid Guid { get; }

    /// <summary>
    /// A source of the in the cache. Can be used to refresh the cache. May not always be
    /// available (e.g. external source)
    /// </summary>
    internal IXLPivotSource Source { get; set; }

    internal String? WorkbookCacheRelId { get; set; }

    internal XLPivotCache AddCachedField(String fieldName, XLPivotCacheValues fieldValues)
    {
        if (this._fieldNames.Contains(fieldName, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException($"Source already contains field {fieldName}.");
        }

        this.AddField(fieldName, fieldValues);
        return this;
    }

    /// <summary>
    /// Try to get a field index for a field name.
    /// </summary>
    /// <param name="fieldName">Name of the field.</param>
    /// <param name="index">The found index, start at 0.</param>
    /// <returns>True if source contains the field.</returns>
    internal bool TryGetFieldIndex(String fieldName, out int index) =>
        this._fieldIndexes.TryGetValue(fieldName, out index);

    internal bool ContainsField(String fieldName) => this._fieldIndexes.ContainsKey(fieldName);

    internal XLPivotCacheValues GetFieldValues(int fieldIndex) => this._values[fieldIndex];

    internal XLPivotCacheSharedItems GetFieldSharedItems(int fieldIndex) =>
        this._values[fieldIndex].SharedItems;

    internal void AllocateRecordCapacity(int recordCount)
    {
        foreach (XLPivotCacheValues fieldValues in this._values)
        {
            fieldValues.AllocateCapacity(recordCount);
        }
    }

    private String AdjustedFieldName(String header)
    {
        string modifiedHeader = header;
        int i = 1;
        while (this._fieldNames.Contains(modifiedHeader, StringComparer.OrdinalIgnoreCase))
        {
            i++;
            modifiedHeader = header + i.ToInvariantString();
        }

        return modifiedHeader;
    }

    private void AddField(String fieldName, XLPivotCacheValues fieldValues)
    {
        this._fieldIndexes.Add(fieldName, this._fieldNames.Count);
        this._fieldNames.Add(fieldName);
        this._values.Add(fieldValues);
    }

    private void SetExcelDefaults() => this.SaveSourceData = true;
}

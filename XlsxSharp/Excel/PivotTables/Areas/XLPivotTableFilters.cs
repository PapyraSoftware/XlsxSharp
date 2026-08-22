// Keep this file CodeMaid organised and cleaned

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XlsxSharp.Excel;

/// <summary>
/// Page/filter fields of a <see cref="XLPivotTable"/>. It determines filter values and layout.
/// It is accessible through fluent API <see cref="XLPivotTable.ReportFilters"/>.
/// </summary>
internal class XLPivotTableFilters : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Filter fields in correct order. The layout is determined by
    /// <see cref="XLPivotTable.FilterFieldsPageWrap"/> and
    /// <see cref="XLPivotTable.FilterAreaOrder"/>.
    /// </summary>
    private readonly List<XLPivotPageField> _fields = [];

    internal XLPivotTableFilters(XLPivotTable pivotTable) => this._pivotTable = pivotTable;

    IXLPivotField IXLPivotFields.Add(string sourceName) => this.Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(string sourceName, string customName) =>
        this.Add(sourceName, customName);

    public void Clear()
    {
        foreach (XLPivotPageField field in this._fields)
        {
            this._pivotTable.RemoveFieldFromAxis(field.Field);
        }

        this._fields.Clear();
    }

    public bool Contains(string sourceName) => this.IndexOf(sourceName) >= 0;

    public bool Contains(IXLPivotField pivotField) => this.Contains(pivotField.SourceName);

    public IXLPivotField Get(string sourceName)
    {
        if (!this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex fieldIndex))
        {
            throw new KeyNotFoundException(
                $"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}."
            );
        }

        XLPivotPageField? filterField = this._fields.SingleOrDefault(f => f.Field == fieldIndex);
        if (filterField is null)
        {
            throw new KeyNotFoundException(
                $"Field with source name '{sourceName}' not found in {XLPivotAxis.AxisPage}."
            );
        }

        return new XLPivotTablePageField(this._pivotTable, filterField);
    }

    public IXLPivotField Get(int index)
    {
        if (index < 0 || index >= this._fields.Count)
        {
            throw new IndexOutOfRangeException();
        }

        return new XLPivotTablePageField(this._pivotTable, this._fields[index]);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IEnumerator<XLPivotTablePageField> GetEnumerator()
    {
        foreach (XLPivotPageField field in this._fields)
        {
            yield return new XLPivotTablePageField(this._pivotTable, field);
        }
    }

    public int IndexOf(string sourceName)
    {
        if (!this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex fieldIndex))
        {
            return -1;
        }

        return this._fields.FindIndex(f => f.Field == fieldIndex);
    }

    public int IndexOf(IXLPivotField pf) => this.IndexOf(pf.SourceName);

    public void Remove(string sourceName)
    {
        int index = this.IndexOf(sourceName);
        if (index == -1)
        {
            return;
        }

        int heightDifference = this.GetHeightDifference(-1);
        Area movedArea = this._pivotTable.Area.ShiftRows(heightDifference);

        this._fields.RemoveAt(index);
        this._pivotTable.RemoveFieldFromAxis(index);

        this._pivotTable.Area = movedArea;
    }

    internal IReadOnlyList<XLPivotPageField> Fields => this._fields;

    internal XLPivotTablePageField Add(string sourceName, string customName)
    {
        if (sourceName == XLConstants.PivotTable.ValuesSentinalLabel)
        {
            throw new ArgumentException(
                nameof(sourceName),
                $"The column '{sourceName}' does not appear in the source range."
            );
        }

        int heightDifference = this.GetHeightDifference(1);
        Area movedArea = this._pivotTable.Area.ShiftRows(heightDifference);

        FieldIndex fieldIndex = this._pivotTable.AddFieldToAxis(
            sourceName,
            customName,
            XLPivotAxis.AxisPage
        );
        XLPivotPageField filterField = new(fieldIndex);
        this._fields.Add(filterField);

        this._pivotTable.Area = movedArea;
        return new XLPivotTablePageField(this._pivotTable, filterField);
    }

    internal bool Contains(FieldIndex fieldIndex) =>
        this._fields.FindIndex(f => f.Field == fieldIndex) >= 0;

    internal void AddField(XLPivotPageField pageField) => this._fields.Add(pageField);

    /// <summary>
    /// Number of rows/cols occupied by the filter area. Filter area is above the pivot table and it
    /// optional (i.e. size <c>0</c> indicates no filter).
    /// </summary>
    internal (int Width, int Height) GetSize() =>
        GetSize(
            this._fields.Count,
            this._pivotTable.FilterAreaOrder,
            this._pivotTable.FilterFieldsPageWrap
        );

    /// <summary>
    /// Number of rows/cols occupied by the filter area, including the gap below, if there is at least one filter.
    /// </summary>
    internal (int Width, int Height) GetSizeWithGap() =>
        GetSizeWithGap(
            this._fields.Count,
            this._pivotTable.FilterAreaOrder,
            this._pivotTable.FilterFieldsPageWrap
        );

    private int GetHeightDifference(int fieldChangeCount)
    {
        int originalHeight = GetSizeWithGap(
            this._fields.Count,
            this._pivotTable.FilterAreaOrder,
            this._pivotTable.FilterFieldsPageWrap
        ).Height;
        int modifiedHeight = GetSizeWithGap(
            this._fields.Count + fieldChangeCount,
            this._pivotTable.FilterAreaOrder,
            this._pivotTable.FilterFieldsPageWrap
        ).Height;
        return modifiedHeight - originalHeight;
    }

    private static (int Width, int Height) GetSize(
        int fieldCount,
        XLFilterAreaOrder order,
        int filterWrap
    )
    {
        if (filterWrap == 0)
        {
            filterWrap = int.MaxValue;
        }

        int dim1 = Math.DivRem(fieldCount, filterWrap, out int dim2);
        dim1 = fieldCount > 0 ? dim1 + 1 : dim1;

        return order switch
        {
            XLFilterAreaOrder.DownThenOver => new ValueTuple<int, int>(dim1, dim2),
            XLFilterAreaOrder.OverThenDown => new ValueTuple<int, int>(dim2, dim1),
            _ => throw new UnreachableException(),
        };
    }

    private static (int Width, int Height) GetSizeWithGap(
        int fieldCount,
        XLFilterAreaOrder order,
        int filterWrap
    )
    {
        (int Width, int Height) filtersSize = GetSize(fieldCount, order, filterWrap);
        return filtersSize.Height > 0 ? (filtersSize.Width, filtersSize.Height + 1) : filtersSize;
    }
}

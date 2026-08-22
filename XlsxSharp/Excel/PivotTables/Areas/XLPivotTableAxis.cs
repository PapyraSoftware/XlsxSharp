using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel;

/// <summary>
/// A description of one axis (<see cref="XLPivotTable.RowAxis"/>/<see cref="XLPivotTable.ColumnAxis"/>)
/// of a <see cref="XLPivotTable"/>. It consists of fields in a specific order and values that make up
/// individual rows/columns of the axis.
/// </summary>
/// <remarks>
/// [ISO-29500] 18.10.1.17 colItems (Column Items), 18.10.1.84 rowItems (Row Items).
/// </remarks>
internal class XLPivotTableAxis : IXLPivotFields
{
    private readonly XLPivotTable _pivotTable;

    private readonly XLPivotAxis _axis;

    /// <summary>
    /// Fields displayed on the axis, in the order of the fields on the axis.
    /// </summary>
    private readonly List<FieldIndex> _fields = [];

    /// <summary>
    /// Values of one row/column in an axis. Items are not kept in sync with <see cref="_fields"/>.
    /// </summary>
    private readonly List<XLPivotFieldAxisItem> _axisItems = [];

    internal XLPivotTableAxis(XLPivotTable pivotTable, XLPivotAxis axis)
    {
        this._pivotTable = pivotTable;
        this._axis = axis;
    }

    /// <summary>
    /// A list of fields to displayed on the axis. It determines which fields and in what order
    /// should the fields be displayed.
    /// </summary>
    internal IReadOnlyList<FieldIndex> Fields => this._fields;

    /// <summary>
    /// Individual row/column parts of the axis.
    /// </summary>
    internal IReadOnlyList<XLPivotFieldAxisItem> Items => this._axisItems;

    internal bool ContainsDataField => this._fields.Any(x => x.IsDataField);

    IXLPivotField IXLPivotFields.Add(String sourceName) => this.Add(sourceName, sourceName);

    IXLPivotField IXLPivotFields.Add(String sourceName, String customName) =>
        this.Add(sourceName, customName);

    void IXLPivotFields.Clear() => this.Clear();

    Boolean IXLPivotFields.Contains(String sourceName) => this.Contains(sourceName);

    Boolean IXLPivotFields.Contains(IXLPivotField pivotField) =>
        this.Contains(pivotField.SourceName);

    IXLPivotField IXLPivotFields.Get(String sourceName)
    {
        if (
            !this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex index)
            || !this._fields.Contains(index)
        )
        {
            throw new KeyNotFoundException(
                $"Field with source name '{sourceName}' not found in {this._axis}."
            );
        }

        return new XLPivotTableAxisField(this._pivotTable, index);
    }

    IXLPivotField IXLPivotFields.Get(Int32 index)
    {
        if (index < 0 || index >= this._fields.Count)
        {
            throw new IndexOutOfRangeException();
        }

        return new XLPivotTableAxisField(this._pivotTable, this._fields[index]);
    }

    Int32 IXLPivotFields.IndexOf(String sourceName) => this.IndexOf(sourceName);

    Int32 IXLPivotFields.IndexOf(IXLPivotField pf) => this.IndexOf(pf.SourceName);

    void IXLPivotFields.Remove(String sourceName)
    {
        int index = this.IndexOf(sourceName);
        if (index == -1)
        {
            return;
        }

        this._pivotTable.RemoveFieldFromAxis(this._fields[index]);
        this._fields.RemoveAt(index);
    }

    IEnumerator<IXLPivotField> IEnumerable<IXLPivotField>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IEnumerator<XLPivotTableAxisField> GetEnumerator()
    {
        foreach (FieldIndex fieldIndex in this._fields)
        {
            yield return new XLPivotTableAxisField(this._pivotTable, fieldIndex);
        }
    }

    internal int IndexOf(FieldIndex index) => this._fields.IndexOf(index);

    internal bool Contains(string sourceName)
    {
        if (!this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex index))
        {
            return false;
        }

        return this._fields.Contains(index);
    }

    /// <summary>
    /// Add field to the axis, as an index.
    /// </summary>
    internal void AddField(FieldIndex fieldIndex)
    {
        if (this._pivotTable.IsFieldUsedOnAxis(fieldIndex))
        {
            throw new ArgumentException("Field is already used on an axis.");
        }

        this._fields.Add(fieldIndex);
    }

    private XLPivotTableAxisField Add(String sourceName, String customName)
    {
        XLPivotTableAxisField field = this.AddField(sourceName, customName);

        // Excel by default adds a subtotal, but previous versions of XlsxSharp didn't have them,
        // so keep API behavior.
        if (field.Offset != FieldIndex.DataField.Value)
        {
            this._pivotTable.PivotFields[field.Offset].RemoveSubtotal(XLSubtotalFunction.Automatic);
        }

        return field;
    }

    internal XLPivotTableAxisField AddField(String sourceName, String customName)
    {
        FieldIndex index = this._pivotTable.AddFieldToAxis(sourceName, customName, this._axis);
        this._fields.Add(index);
        return new XLPivotTableAxisField(this._pivotTable, index);
    }

    /// <summary>
    /// Add a row/column axis values (i.e. values visible on the axis).
    /// </summary>
    internal void AddItem(XLPivotFieldAxisItem axisItem) => this._axisItems.Add(axisItem);

    internal void Clear()
    {
        foreach (FieldIndex fieldIndex in this._fields)
        {
            this._pivotTable.RemoveFieldFromAxis(fieldIndex);
        }

        this._axisItems.Clear();
        this._fields.Clear();
    }

    private Int32 IndexOf(String sourceName)
    {
        if (!this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex fieldIndex))
        {
            return -1;
        }

        return this._fields.IndexOf(fieldIndex);
    }
}

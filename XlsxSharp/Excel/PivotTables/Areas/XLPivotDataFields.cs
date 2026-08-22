using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.PivotValues;

namespace XlsxSharp.Excel;

/// <summary>
/// A collection of <see cref="XLPivotDataField"/>.
/// </summary>
internal class XLPivotDataFields : IXLPivotValues, IReadOnlyCollection<XLPivotDataField>
{
    private readonly XLPivotTable _pivotTable;

    /// <summary>
    /// Fields displayed in the data area of the pivot table, in the order fields are displayed.
    /// </summary>
    private readonly List<XLPivotDataField> _fields = [];

    internal XLPivotDataFields(XLPivotTable pivotTable)
    {
        this._pivotTable = pivotTable;
    }

    public int Count => this._fields.Count;

    #region IXLPivotValues

    public IXLPivotValue Add(string sourceName)
    {
        return this.AddField(sourceName, sourceName);
    }

    public IXLPivotValue Add(string sourceName, string customName)
    {
        return this.AddField(sourceName, customName);
    }

    public void Clear()
    {
        this._fields.Clear();
        foreach (XLPivotDataField field in this._fields)
        {
            this._pivotTable.RemoveFieldFromAxis(field.Field);
        }
    }

    public bool Contains(string customName)
    {
        return this.IndexOf(customName) != -1;
    }

    public bool Contains(IXLPivotValue pivotValue)
    {
        return this.Contains(pivotValue.CustomName);
    }

    public IXLPivotValue Get(string customName)
    {
        XLPivotDataField? dataField = this._fields.SingleOrDefault(x =>
            XlsxSharp.XLHelper.NameComparer.Equals(x.CustomName, customName)
        );
        if (dataField is null)
        {
            throw new KeyNotFoundException($"Unable to find data field for '{customName}'.");
        }

        return dataField;
    }

    public IXLPivotValue Get(int index)
    {
        return this._fields[index];
    }

    public int IndexOf(string customName)
    {
        return this._fields.FindIndex(x =>
            XlsxSharp.XLHelper.NameComparer.Equals(x.CustomName, customName)
        );
    }

    public int IndexOf(IXLPivotValue pivotValue)
    {
        return this.IndexOf(pivotValue.CustomName);
    }

    public void Remove(string customName)
    {
        int index = this.IndexOf(customName);
        if (index == -1)
        {
            return;
        }

        XLPivotDataField dataField = this._fields[index];
        this._pivotTable.RemoveFieldFromAxis(dataField.Field);
        this._fields.Remove(dataField);
    }

    IEnumerator<IXLPivotValue> IEnumerable<IXLPivotValue>.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    #endregion

    internal XLPivotDataField AddField(string sourceName, string? customName)
    {
        if (!this._pivotTable.TryGetSourceNameFieldIndex(sourceName, out FieldIndex fieldIndex))
        {
            string validNames = string.Join("','", this._pivotTable.PivotCache.FieldNames);
            throw new ArgumentOutOfRangeException(
                nameof(sourceName),
                $"Field '{sourceName}' is not in the fields of a pivot cache. Should be one of '{validNames}'."
            );
        }

        if (fieldIndex.IsDataField)
        {
            throw new ArgumentException("'Values' field can be used only on row or column axis.");
        }

        XLPivotDataField dataField = new(this._pivotTable, fieldIndex.Value)
        {
            DataFieldName = customName,
        };
        this.AddField(dataField);

        // If there are multiple values, at least axis must contain 'data' field.
        // Otherwise, Excel requires a repair.
        if (
            this._fields.Count > 1
            && !this._pivotTable.RowAxis.ContainsDataField
            && !this._pivotTable.ColumnAxis.ContainsDataField
        )
        {
            this._pivotTable.ColumnLabels.Add(XLConstants.PivotTable.ValuesSentinalLabel);
        }

        return dataField;
    }

    internal void AddField(XLPivotDataField dataField)
    {
        // Excel invariant - data field must have the flag if and only if it is in the data fields collection.
        this._fields.Add(dataField);
        this._pivotTable.PivotFields[dataField.Field].DataField = true;
    }

    public IEnumerator<XLPivotDataField> GetEnumerator()
    {
        return this._fields.GetEnumerator();
    }
}

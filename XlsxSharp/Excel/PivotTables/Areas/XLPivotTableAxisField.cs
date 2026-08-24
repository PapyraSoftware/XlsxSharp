#nullable disable
using XlsxSharp.Excel.PivotStyleFormats;

namespace XlsxSharp.Excel;

/// <summary>
/// A fluent API for one field in <see cref="XLPivotTableAxis"/>, either
/// <see cref="XLPivotTable.RowLabels"/> or <see cref="XLPivotTable.ColumnLabels"/>.
/// </summary>
internal class XLPivotTableAxisField : IXLPivotField
{
    private readonly XLPivotTable _pivotTable;
    private readonly FieldIndex _index;

    internal XLPivotTableAxisField(XLPivotTable pivotTable, FieldIndex index)
    {
        this._pivotTable = pivotTable;
        this._index = index;
    }

    #region IXLPivotField memebers

    public string SourceName
    {
        get
        {
            if (this._index.IsDataField)
            {
                return XLConstants.PivotTable.ValuesSentinalLabel;
            }

            return this._pivotTable.PivotCache.FieldNames[this._index];
        }
    }

    public string CustomName
    {
        get => this.GetFieldValue(f => f.Name, this._pivotTable.DataCaption);
        set
        {
            if (this._index.IsDataField)
            {
                this._pivotTable.DataCaption = value;
                return;
            }

            if (
                this._pivotTable.TryGetCustomNameFieldIndex(value, out FieldIndex idx)
                && idx != this._index
            )
            {
                throw new ArgumentException(
                    $"Custom name '{value}' is already used by another field."
                );
            }

            this._pivotTable.PivotFields[this._index].Name = value;
        }
    }

    public string SubtotalCaption
    {
        get => this.GetFieldValue(f => f.SubtotalCaption, string.Empty);
        set => this.GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XLSubtotalFunction> Subtotals
    {
        get
        {
            HashSet<XLSubtotalFunction> subtotal = this.GetField().Subtotals;
            bool isCustomSubtotal =
                subtotal.Count > 1
                || (subtotal.Count > 0 && !subtotal.Contains(XLSubtotalFunction.Automatic));
            if (isCustomSubtotal)
            {
                // When subtotal is custom, the automatic is not shown
                subtotal = [.. subtotal];
                subtotal.Remove(XLSubtotalFunction.Automatic);
            }

            return subtotal;
        }
    }

    public bool IncludeNewItemsInFilter
    {
        get => this.GetFieldValue(f => f.IncludeNewItemsInFilter, false);
        set => this.GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => this.GetFieldValue(f => f.Outline, true);
        set => this.GetField().Outline = value;
    }
    public bool Compact
    {
        get => this.GetFieldValue(f => f.Compact, true);
        set => this.GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => this.GetFieldValue(f => f.SubtotalTop, true);
        set => this.GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => this.GetFieldValue(f => f.RepeatItemLabels, false);
        set => this.GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => this.GetFieldValue(f => f.InsertBlankRow, false);
        set => this.GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => this.GetFieldValue(f => f.ShowAll, true);
        set => this.GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => this.GetFieldValue(f => f.InsertPageBreak, false);
        set => this.GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get => this.GetFieldValue(f => !f.Items.Any(i => i.ShowDetails), false);
        set
        {
            foreach (XLPivotFieldItem item in this.GetField().Items)
            {
                item.ShowDetails = !value;
            }
        }
    }

    public XLPivotSortType SortType
    {
        get => this.GetFieldValue(f => f.SortType, XLPivotSortType.Default);
        set => this.GetField().SortType = value;
    }

    public IReadOnlyList<XLCellValue> SelectedValues => Array.Empty<XLCellValue>();

    public IXLPivotFieldStyleFormats StyleFormats =>
        new XLPivotTableAxisFieldStyleFormats(this._pivotTable, this);

    public bool IsOnRowAxis =>
        this.GetFieldValue(f => f.Axis == XLPivotAxis.AxisRow, this._pivotTable.DataOnRows);

    public bool IsOnColumnAxis =>
        this.GetFieldValue(f => f.Axis == XLPivotAxis.AxisCol, !this._pivotTable.DataOnRows);

    public bool IsInFilterList => false;

    public int Offset => this._index;

    public IXLPivotField SetCustomName(string value)
    {
        this.CustomName = value;
        return this;
    }

    public IXLPivotField SetSubtotalCaption(string value)
    {
        this.SubtotalCaption = value;
        return this;
    }

    public IXLPivotField AddSubtotal(XLSubtotalFunction value)
    {
        this.GetField().AddSubtotal(value);
        return this;
    }

    public IXLPivotField SetIncludeNewItemsInFilter(bool value)
    {
        this.IncludeNewItemsInFilter = value;
        return this;
    }

    public IXLPivotField SetLayout(XLPivotLayout value)
    {
        this.GetField().SetLayout(value);
        return this;
    }

    public IXLPivotField SetSubtotalsAtTop(bool value)
    {
        this.SubtotalsAtTop = value;
        return this;
    }

    public IXLPivotField SetRepeatItemLabels(bool value)
    {
        this.RepeatItemLabels = value;
        return this;
    }

    public IXLPivotField SetInsertBlankLines(bool value)
    {
        this.InsertBlankLines = value;
        return this;
    }

    public IXLPivotField SetShowBlankItems(bool value)
    {
        this.ShowBlankItems = value;
        return this;
    }

    public IXLPivotField SetInsertPageBreaks(bool value)
    {
        this.InsertPageBreaks = value;
        return this;
    }

    public IXLPivotField SetCollapsed(bool value)
    {
        this.Collapsed = true;
        return this;
    }

    public IXLPivotField SetSort(XLPivotSortType value)
    {
        this.SortType = value;
        return this;
    }

    public IXLPivotField AddSelectedValue(XLCellValue value) => this;

    public IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values) => this;

    #endregion IXLPivotField members

    internal XLPivotAxis Axis => this.IsOnColumnAxis ? XLPivotAxis.AxisCol : XLPivotAxis.AxisRow;

    /// <summary>
    /// Get position of the field on the axis, starting at 0.
    /// </summary>
    internal int Position
    {
        get
        {
            XLPivotTableAxis axis = this.IsOnColumnAxis
                ? this._pivotTable.ColumnAxis
                : this._pivotTable.RowAxis;
            int position = axis.IndexOf(this._index);
            if (position == -1)
            {
                throw new InvalidOperationException("Field is not on the axis.");
            }

            return position;
        }
    }

    private XLPivotTableField GetField()
    {
        if (this._index.IsDataField)
        {
            throw new InvalidOperationException("Can't set this property on a data field.");
        }

        return this._pivotTable.PivotFields[this._index];
    }

    private T GetFieldValue<T>(Func<XLPivotTableField, T> getter, T dataFieldValue)
    {
        if (this._index.IsDataField)
        {
            return dataFieldValue;
        }

        XLPivotTableField field = this._pivotTable.PivotFields[this._index];
        return getter(field);
    }
}

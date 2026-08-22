#nullable disable
using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.PivotStyleFormats;

namespace XlsxSharp.Excel;

/// <summary>
/// Fluent API for filter fields of a <see cref="XLPivotTable"/>. This class shouldn't contain any
/// state, only logic to change state per API.
/// </summary>
internal class XLPivotTablePageField : IXLPivotField
{
    private readonly XLPivotTable _pivotTable;
    private readonly XLPivotPageField _filterField;

    internal XLPivotTablePageField(XLPivotTable pivotTable, XLPivotPageField filterField)
    {
        this._pivotTable = pivotTable;
        this._filterField = filterField;
    }

    public string SourceName => this._pivotTable.PivotCache.FieldNames[this._filterField.Field];

    public string CustomName
    {
        get => this.GetField().Name;
        set => this.GetField().Name = value;
    }

    public string SubtotalCaption
    {
        get => this.GetField().SubtotalCaption;
        set => this.GetField().SubtotalCaption = value;
    }

    public IReadOnlyCollection<XLSubtotalFunction> Subtotals => this.GetField().Subtotals;

    public bool IncludeNewItemsInFilter
    {
        get => this.GetField().IncludeNewItemsInFilter;
        set => this.GetField().IncludeNewItemsInFilter = value;
    }

    public bool Outline
    {
        get => this.GetField().Outline;
        set => this.GetField().Outline = value;
    }

    public bool Compact
    {
        get => this.GetField().Compact;
        set => this.GetField().Compact = value;
    }

    public bool? SubtotalsAtTop
    {
        get => this.GetField().SubtotalTop;
        set => this.GetField().SubtotalTop = value ?? true;
    }

    public bool RepeatItemLabels
    {
        get => this.GetField().RepeatItemLabels;
        set => this.GetField().RepeatItemLabels = value;
    }

    public bool InsertBlankLines
    {
        get => this.GetField().InsertBlankRow;
        set => this.GetField().InsertBlankRow = value;
    }

    public bool ShowBlankItems
    {
        get => this.GetField().ShowAll;
        set => this.GetField().ShowAll = value;
    }

    public bool InsertPageBreaks
    {
        get => this.GetField().InsertPageBreak;
        set => this.GetField().InsertPageBreak = value;
    }

    public bool Collapsed
    {
        get => this.GetField().Collapsed;
        set => this.GetField().Collapsed = value;
    }

    public XLPivotSortType SortType
    {
        get => this.GetField().SortType;
        set => this.GetField().SortType = value;
    }

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
        this.Collapsed = value;
        return this;
    }

    public IXLPivotField SetSort(XLPivotSortType value)
    {
        this.SortType = value;
        return this;
    }

    public IReadOnlyList<XLCellValue> SelectedValues
    {
        get
        {
            IEnumerable<XLPivotFieldItem> shownItems = this.GetField().Items.Where(i => !i.Hidden);
            List<XLCellValue> selectedValues = [];
            foreach (XLPivotFieldItem selectedItem in shownItems)
            {
                XLCellValue? selectedValue = selectedItem.GetValue();
                if (selectedValue is not null)
                {
                    selectedValues.Add(selectedValue.Value);
                }
            }

            return selectedValues;
        }
    }

    public IXLPivotField AddSelectedValue(XLCellValue value)
    {
        // Try to keep the original behavior of XlsxSharp - it always allows multiple selected items for added values.
        // But it's complete kludge with no sane semantic that will be nuked ASAP.
        XLPivotTableField pivotField = this.GetField();

        bool nothingSelected =
            this._filterField.ItemIndex is null && !pivotField.MultipleItemSelectionAllowed;
        if (nothingSelected)
        {
            XLPivotFieldItem fieldItem = pivotField.GetOrAddItem(value);
            this._filterField.ItemIndex = fieldItem.ItemIndex;
            return this;
        }

        bool oneItemSelected =
            this._filterField.ItemIndex is not null && !pivotField.MultipleItemSelectionAllowed;
        if (oneItemSelected)
        {
            // Switch to multiple
            pivotField.MultipleItemSelectionAllowed = true;
            foreach (
                XLPivotFieldItem item in pivotField.Items.Where(x =>
                    x.ItemType == XLPivotItemType.Data
                )
            )
            {
                item.Hidden = true;
            }

            XLPivotFieldItem selectedItem = pivotField.Items.Single(i =>
                i.ItemIndex == this._filterField.ItemIndex
            );
            selectedItem.Hidden = false;
            this._filterField.ItemIndex = null;
            XLPivotFieldItem fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
        else
        {
            // Add another item to selected item filters.
            XLPivotFieldItem fieldItem = pivotField.GetOrAddItem(value);
            fieldItem.Hidden = false;
            return this;
        }
    }

    public IXLPivotField AddSelectedValues(IEnumerable<XLCellValue> values)
    {
        foreach (XLCellValue value in values)
        {
            this.AddSelectedValue(value);
        }

        return this;
    }

    public IXLPivotFieldStyleFormats StyleFormats =>
        throw new NotImplementedException("Styles for filter fields are not yet implemented.");
    public bool IsOnRowAxis => false;
    public bool IsOnColumnAxis => false;
    public bool IsInFilterList => true;
    public int Offset => this._filterField.Field;

    private XLPivotTableField GetField()
    {
        return this._pivotTable.PivotFields[this._filterField.Field];
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.PivotValues;

namespace XlsxSharp.Excel;

/// <summary>
/// A field that describes calculation of value to display in the <see cref="XLPivotAreaType.Data"/>
/// area of pivot table.
/// </summary>
internal class XLPivotDataField : IXLPivotValue
{
    private const int BaseFieldDefaultValue = -1;
    private const int BaseItemPreviousValue = 1048828;
    private const int BaseItemNextValue = 1048829;
    private const int BaseItemDefaultValue = 1048832;

    private readonly XLPivotTable _pivotTable;

    private XLNumberFormat? _numberFormat;
    private int _baseField = BaseFieldDefaultValue;
    private uint _baseItem = BaseItemDefaultValue;
    private XLPivotCalculation _showDataAsFormat = XLPivotCalculation.Normal;
    private XLPivotSummary _subtotal = XLPivotSummary.Sum;

    internal XLPivotDataField(XLPivotTable pivotTable, int field)
    {
        this._pivotTable = pivotTable;
        this.Field = field;
    }

    /// <summary>
    /// Custom name of the data field (e.g. <em>Sum of Sold</em>). Can be left empty to keep same
    /// as source name. Use <see cref="CustomName"/> to get value with fallback.
    /// </summary>
    /// <remarks>
    /// For data fields, the name is duplicated at <see cref="XLPivotTableField.Name"/> and here.
    /// This property has a preference.
    /// </remarks>
    internal string? DataFieldName { get; set; }

    /// <summary>
    /// Field index to <see cref="XLPivotTable.PivotFields"/>.
    /// </summary>
    /// <remarks>
    /// Unlike axis, this field index can't be <c>-2</c> for data fields. That field can't be in
    /// the data area.
    /// </remarks>
    internal int Field { get; }

    /// <summary>
    /// An aggregation function that calculates the value to display in the data cells of pivot area.
    /// </summary>
    public XLPivotSummary Subtotal
    {
        get => this._subtotal;
        set => this._subtotal = value;
    }

    /// <summary>
    /// A calculation takes value calculated by <see cref="Subtotal"/> aggregation and transforms
    /// it into the final value to display to the user. The calculation might need
    /// <see cref="BaseField"/> and/or <see cref="BaseItem"/>.
    /// </summary>
    public XLPivotCalculation ShowDataAsFormat
    {
        get => this._showDataAsFormat;
        init => this._showDataAsFormat = value;
    }

    /// <summary>
    /// Index to the base field (<see cref="XLPivotTable.PivotFields"/>) when
    /// <see cref="ShowDataAsFormat"/> needs a field for its calculation.
    /// </summary>
    public int BaseField
    {
        get => this._baseField;
        init => this._baseField = value;
    }

    /// <summary>
    /// Index to the base item of <see cref="BaseField"/> when <see cref="ShowDataAsFormat"/> needs
    /// an item for its calculation.
    /// </summary>
    public uint BaseItem
    {
        get => this._baseItem;
        init => this._baseItem = value;
    }

    /// <summary>
    /// Formatting to apply to the data field. If <see cref="XLPivotFormat"/> disagree, this has precedence.
    /// </summary>
    internal XLNumberFormat? NumberFormatValue
    {
        get => this._numberFormat;
        set
        {
            this._numberFormat = value is not null
                ? this._pivotTable.Worksheet.Workbook.Styles.RegisterNumberFormat(value.Value)
                : null;
        }
    }

    #region IXLPivotValue

    public string? BaseFieldName
    {
        get
        {
            IReadOnlyList<string> sourceNames = this._pivotTable.PivotCache.FieldNames;
            if (this._baseField < 0 || this._baseField >= sourceNames.Count)
            {
                return null;
            }

            return sourceNames[this._baseField];
        }
        set
        {
            if (value is null)
            {
                this._baseField = BaseFieldDefaultValue;
                return;
            }

            if (!this._pivotTable.TryGetSourceNameFieldIndex(value, out FieldIndex index))
            {
                throw new ArgumentOutOfRangeException($"Source name '{value}' not found.");
            }

            this._baseField = index;
        }
    }

    public XLCellValue BaseItemValue
    {
        get
        {
            bool baseFieldSpecified = this._baseField != BaseFieldDefaultValue;
            if (!baseFieldSpecified)
            {
                return Blank.Value;
            }

            bool baseItemSpecified = this._baseItem != BaseItemDefaultValue;
            if (!baseItemSpecified)
            {
                return Blank.Value;
            }

            if (this._baseItem == BaseItemPreviousValue)
            {
                return Blank.Value;
            }

            if (this._baseItem == BaseItemNextValue)
            {
                return Blank.Value;
            }

            XLPivotTableField baseField = this._pivotTable.PivotFields[this._baseField];
            XLPivotFieldItem fieldItem = baseField.Items[checked((int)this.BaseItem)];
            return fieldItem.GetValue() ?? Blank.Value;
        }
        set
        {
            if (this._baseField == BaseItemDefaultValue)
            {
                throw new InvalidOperationException("Base field not specified for the field.");
            }

            XLPivotTableField pivotField = this._pivotTable.PivotFields[this._baseField];
            XLPivotFieldItem fieldItem = pivotField.GetOrAddItem(value);
            int itemIndex = fieldItem.ItemIndex ?? BaseFieldDefaultValue;
            this._baseItem = checked((uint)itemIndex);
        }
    }

    public XLPivotCalculation Calculation
    {
        get => this.ShowDataAsFormat;
        set => this._showDataAsFormat = value;
    }

    public XLPivotCalculationItem CalculationItem
    {
        get
        {
            return this._baseItem switch
            {
                BaseItemPreviousValue => XLPivotCalculationItem.Previous,
                BaseItemNextValue => XLPivotCalculationItem.Next,
                _ => XLPivotCalculationItem.Value,
            };
        }
        set
        {
            switch (value)
            {
                case XLPivotCalculationItem.Previous:
                    this._baseItem = BaseItemPreviousValue;
                    break;
                case XLPivotCalculationItem.Next:
                    this._baseItem = BaseItemNextValue;
                    break;
                case XLPivotCalculationItem.Value:
                    // Calculation value should be set in tandem with the base item value.
                    // Base item other than prev/next special constants is implicitly a value.
                    if (this.BaseItem is BaseItemPreviousValue or BaseItemNextValue)
                    {
                        // If value is not yet set, just use unspecified value. User should
                        // set value by calling `BaseItemValue` after calling this, but Excel
                        // accepts valid base field with unspecified item without need to repair.
                        this._baseItem = BaseItemDefaultValue;
                    }

                    // When base item is not a valid reference to the field.Items, Excel
                    // tries to repair the workbook, so user should always set base value.
                    break;
                default:
                    throw new UnreachableException();
            }
        }
    }

    public string CustomName
    {
        get =>
            this.DataFieldName
            ?? this._pivotTable.PivotFields[this.Field].Name
            ?? this._pivotTable.PivotCache.FieldNames[this.Field];
        set => this.DataFieldName = value;
    }

    public IXLPivotValueFormat NumberFormat => new XLPivotValueFormat(this);

    public string SourceName => this._pivotTable.PivotCache.FieldNames[this.Field];

    public XLPivotSummary SummaryFormula
    {
        get => this.Subtotal;
        set => this._subtotal = value;
    }

    public IXLPivotValue SetBaseFieldName(string value)
    {
        this.BaseFieldName = value;
        return this;
    }

    public IXLPivotValue SetBaseItemValue(XLCellValue value)
    {
        this.BaseItemValue = value;
        return this;
    }

    public IXLPivotValue SetCalculation(XLPivotCalculation value)
    {
        this.Calculation = value;
        return this;
    }

    public IXLPivotValue SetCalculationItem(XLPivotCalculationItem value)
    {
        this.CalculationItem = value;
        return this;
    }

    public IXLPivotValue SetSummaryFormula(XLPivotSummary value)
    {
        this.SummaryFormula = value;
        return this;
    }

    public IXLPivotValueCombination ShowAsDifferenceFrom(string fieldSourceName)
    {
        this.BaseFieldName = fieldSourceName;
        this.SetCalculation(XLPivotCalculation.DifferenceFrom);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValue ShowAsIndex()
    {
        return this.SetCalculation(XLPivotCalculation.Index);
    }

    public IXLPivotValue ShowAsNormal()
    {
        return this.SetCalculation(XLPivotCalculation.Normal);
    }

    public IXLPivotValueCombination ShowAsPercentageDifferenceFrom(string fieldSourceName)
    {
        this.BaseFieldName = fieldSourceName;
        this.SetCalculation(XLPivotCalculation.PercentageDifferenceFrom);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValueCombination ShowAsPercentageFrom(string fieldSourceName)
    {
        this.BaseFieldName = fieldSourceName;
        this.SetCalculation(XLPivotCalculation.PercentageOf);
        return new XLPivotValueCombination(this);
    }

    public IXLPivotValue ShowAsPercentageOfColumn()
    {
        return this.SetCalculation(XLPivotCalculation.PercentageOfColumn);
    }

    public IXLPivotValue ShowAsPercentageOfRow()
    {
        return this.SetCalculation(XLPivotCalculation.PercentageOfRow);
    }

    public IXLPivotValue ShowAsPercentageOfTotal()
    {
        return this.SetCalculation(XLPivotCalculation.PercentageOfTotal);
    }

    public IXLPivotValue ShowAsRunningTotalIn(string fieldSourceName)
    {
        this.BaseFieldName = fieldSourceName;
        return this.SetCalculation(XLPivotCalculation.RunningTotal);
    }

    #endregion IXPivotValue
}

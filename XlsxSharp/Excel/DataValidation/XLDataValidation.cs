using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XlsxSharp.Excel.DataValidation;

internal class XLDataValidation : IXLDataValidation
{
    private readonly XLWorksheet _worksheet;

    internal XLDataValidation(XLWorksheet worksheet)
    {
        this._worksheet = worksheet;
        this.Initialize();
    }

    internal XLWorksheet Worksheet => this._worksheet;

    internal XLAreaList Areas { get; set; } = XLAreaList.Empty;

    public void Clear()
    {
        this.Initialize();
    }

    internal void CopyFrom(IXLDataValidation dataValidation)
    {
        if (dataValidation == this)
        {
            return;
        }

        if (this.Areas.Count == 0)
        {
            this.AddRanges(dataValidation.Ranges);
        }

        this.IgnoreBlanks = dataValidation.IgnoreBlanks;
        this.InCellDropdown = dataValidation.InCellDropdown;
        this.ShowErrorMessage = dataValidation.ShowErrorMessage;
        this.ShowInputMessage = dataValidation.ShowInputMessage;
        this.InputTitle = dataValidation.InputTitle;
        this.InputMessage = dataValidation.InputMessage;
        this.ErrorTitle = dataValidation.ErrorTitle;
        this.ErrorMessage = dataValidation.ErrorMessage;
        this.ErrorStyle = dataValidation.ErrorStyle;
        this.AllowedValues = dataValidation.AllowedValues;
        this.Operator = dataValidation.Operator;
        this.MinValue = dataValidation.MinValue;
        this.MaxValue = dataValidation.MaxValue;
    }

    public Boolean IsDirty()
    {
        return this.AllowedValues != XLAllowedValues.AnyValue
            || (
                this.ShowInputMessage
                && (
                    !String.IsNullOrWhiteSpace(this.InputTitle)
                    || !String.IsNullOrWhiteSpace(this.InputMessage)
                )
            )
            || (
                this.ShowErrorMessage
                && (
                    !String.IsNullOrWhiteSpace(this.ErrorTitle)
                    || !String.IsNullOrWhiteSpace(this.ErrorMessage)
                )
            );
    }

    [MemberNotNull(
        nameof(minValue),
        nameof(maxValue),
        nameof(InputTitle),
        nameof(InputMessage),
        nameof(ErrorTitle),
        nameof(ErrorMessage)
    )]
    private void Initialize()
    {
        this.AllowedValues = XLAllowedValues.AnyValue;
        this.IgnoreBlanks = true;
        this.ShowErrorMessage = true;
        this.ShowInputMessage = true;
        this.InCellDropdown = true;
        this.InputTitle = String.Empty;
        this.InputMessage = String.Empty;
        this.ErrorTitle = String.Empty;
        this.ErrorMessage = String.Empty;
        this.ErrorStyle = XLErrorStyle.Stop;
        this.Operator = XLOperator.Between;
        this.Value = String.Empty;
        this.minValue = String.Empty;
        this.maxValue = String.Empty;
    }

    #region IXLDataValidation Members

    private String maxValue;
    private String minValue;
    public XLAllowedValues AllowedValues { get; set; }

    public XLDateCriteria Date
    {
        get
        {
            this.AllowedValues = XLAllowedValues.Date;
            return new XLDateCriteria(this);
        }
    }

    public XLDecimalCriteria Decimal
    {
        get
        {
            this.AllowedValues = XLAllowedValues.Decimal;
            return new XLDecimalCriteria(this);
        }
    }

    public String ErrorMessage { get; set; }
    public XLErrorStyle ErrorStyle { get; set; }
    public String ErrorTitle { get; set; }
    public Boolean IgnoreBlanks { get; set; }
    public Boolean InCellDropdown { get; set; }
    public String InputMessage { get; set; }
    public String InputTitle { get; set; }
    public String MaxValue
    {
        get => this.maxValue;
        set
        {
            Validate(value);
            this.maxValue = value;
        }
    }
    public String MinValue
    {
        get => this.minValue;
        set
        {
            Validate(value);
            this.minValue = value;
        }
    }
    public XLOperator Operator { get; set; }

    public IEnumerable<IXLRange> Ranges
    {
        get
        {
            XLRanges ranges = new(this.Worksheet);
            foreach (Area area in this.Areas)
            {
                XLRange range = this._worksheet.Range(
                    XLRangeAddress.FromSheetRange(this._worksheet, area)
                );
                ranges.Add(range);
            }

            return ranges;
        }
    }

    public Boolean ShowErrorMessage { get; set; }

    public Boolean ShowInputMessage { get; set; }

    public XLTextLengthCriteria TextLength
    {
        get
        {
            this.AllowedValues = XLAllowedValues.TextLength;
            return new XLTextLengthCriteria(this);
        }
    }

    public XLTimeCriteria Time
    {
        get
        {
            this.AllowedValues = XLAllowedValues.Time;
            return new XLTimeCriteria(this);
        }
    }

    public String Value
    {
        get { return this.MinValue; }
        set { this.MinValue = value; }
    }

    public XLWholeNumberCriteria WholeNumber
    {
        get
        {
            this.AllowedValues = XLAllowedValues.WholeNumber;
            return new XLWholeNumberCriteria(this);
        }
    }

    /// <summary>
    /// Add a range to the collection of ranges this rule applies to.
    /// If the specified range does not belong to the worksheet of the data validation
    /// rule it is transferred to the target worksheet.
    /// </summary>
    /// <param name="range">A range to add.</param>
    public void AddRange(IXLRange range)
    {
        ArgumentNullException.ThrowIfNull(range);

        // Do not add area if the DV has been detached (e.g. consolidation).
        bool isDetached = !this._worksheet.DataValidations.Contains(this);
        if (isDetached)
        {
            return;
        }

        // Ignore sheet of a range
        Area area = Area.FromRangeAddress(range.RangeAddress);
        this._worksheet.DataValidations.AddArea(this, area);
    }

    /// <summary>
    /// Add a collection of ranges to the collection of ranges this rule applies to.
    /// Ranges that do not belong to the worksheet of the data validation
    /// rule are transferred to the target worksheet.
    /// </summary>
    /// <param name="ranges">Ranges to add.</param>
    public void AddRanges(IEnumerable<IXLRange> ranges)
    {
        ranges = ranges ?? Enumerable.Empty<IXLRange>();

        foreach (IXLRange range in ranges)
        {
            this.AddRange(range);
        }
    }

    /// <summary>
    /// Detach data validation rule of all ranges it applies to.
    /// </summary>
    public void ClearRanges()
    {
        this.Areas = XLAreaList.Empty;
    }

    public void Custom(String customValidation)
    {
        this.AllowedValues = XLAllowedValues.Custom;
        this.Value = customValidation;
    }

    public void List(String list)
    {
        this.List(list, true);
    }

    public void List(String list, Boolean inCellDropdown)
    {
        this.AllowedValues = XLAllowedValues.List;
        this.InCellDropdown = inCellDropdown;
        this.Value = list;
    }

    public void List(IXLRange range)
    {
        this.List(range, true);
    }

    public void List(IXLRange range, Boolean inCellDropdown)
    {
        this.List(range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true));
    }

    /// <summary>
    /// Remove the specified range from the collection of range this rule applies to.
    /// </summary>
    /// <param name="range">A range to remove.</param>
    public bool RemoveRange(IXLRange range)
    {
        if (range == null)
        {
            return false;
        }

        Area areaToDelete = SheetArea.From(range).Area;
        XLAreaList originalAreas = this.Areas;
        this.Areas = this.Areas.Without(areaToDelete);
        bool deleted = originalAreas.Count > this.Areas.Count;
        return deleted;
    }

    #endregion IXLDataValidation Members

    private static void Validate(String value)
    {
        if (value.Length > 255)
        {
            throw new ArgumentOutOfRangeException(
                nameof(value),
                "The maximum allowed length of the value is 255 characters."
            );
        }
    }
}

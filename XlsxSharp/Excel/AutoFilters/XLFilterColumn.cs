using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel;

internal class XLFilterColumn : IXLFilterColumn, IXLFilteredColumn, IEnumerable<XLFilter>
{
    private readonly XLAutoFilter _autoFilter;
    private readonly Int32 _column;
    private readonly List<XLFilter> _filters = [];

    public XLFilterColumn(XLAutoFilter autoFilter, Int32 column)
    {
        this._autoFilter = autoFilter;
        this._column = column;
    }

    #region IXLFilterColumn Members

    public void Clear(bool reapply)
    {
        this._filters.Clear();
        this.FilterType = XLFilterType.None;
        if (reapply)
        {
            this._autoFilter.Reapply();
        }
    }

    public IXLFilteredColumn AddFilter(XLCellValue value, bool reapply)
    {
        this.SwitchFilter(XLFilterType.Regular);
        this.AddFilter(XLFilter.CreateRegularFilter(value.ToString()), reapply);
        return this;
    }

    public IXLFilteredColumn AddDateGroupFilter(
        DateTime date,
        XLDateTimeGrouping dateTimeGrouping,
        bool reapply
    )
    {
        this.SwitchFilter(XLFilterType.Regular);
        this.AddFilter(XLFilter.CreateDateGroupFilter(date, dateTimeGrouping), reapply);
        return this;
    }

    public void Top(Int32 value, XLTopBottomType type, bool reapply) =>
        this.SetTopBottom(value, type, takeTop: true, reapply);

    public void Bottom(Int32 value, XLTopBottomType type, bool reapply) =>
        this.SetTopBottom(value, type, takeTop: false, reapply);

    public void AboveAverage(bool reapply) => this.SetAverage(aboveAverage: true, reapply);

    public void BelowAverage(bool reapply) => this.SetAverage(aboveAverage: false, reapply);

    public IXLFilterConnector EqualTo(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value.ToString(), true, reapply);

    public IXLFilterConnector NotEqualTo(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value.ToString(), false, reapply);

    public IXLFilterConnector GreaterThan(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value, XLFilterOperator.GreaterThan, reapply);

    public IXLFilterConnector LessThan(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value, XLFilterOperator.LessThan, reapply);

    public IXLFilterConnector EqualOrGreaterThan(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);

    public IXLFilterConnector EqualOrLessThan(XLCellValue value, Boolean reapply) =>
        this.AddCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);

    public void Between(XLCellValue minValue, XLCellValue maxValue, Boolean reapply) =>
        this.EqualOrGreaterThan(minValue, false).And.EqualOrLessThan(maxValue, reapply);

    public void NotBetween(XLCellValue minValue, XLCellValue maxValue, Boolean reapply) =>
        this.LessThan(minValue, false).Or.GreaterThan(maxValue, reapply);

    public IXLFilterConnector BeginsWith(String value, Boolean reapply) =>
        this.AddCustomFilter(value + "*", true, reapply);

    public IXLFilterConnector NotBeginsWith(String value, Boolean reapply) =>
        this.AddCustomFilter(value + "*", false, reapply);

    public IXLFilterConnector EndsWith(String value, Boolean reapply) =>
        this.AddCustomFilter("*" + value, true, reapply);

    public IXLFilterConnector NotEndsWith(String value, Boolean reapply) =>
        this.AddCustomFilter("*" + value, false, reapply);

    public IXLFilterConnector Contains(String value, Boolean reapply) =>
        this.AddCustomFilter("*" + value + "*", true, reapply);

    public IXLFilterConnector NotContains(String value, Boolean reapply) =>
        this.AddCustomFilter("*" + value + "*", false, reapply);

    public XLFilterType FilterType { get; set; }

    public Int32 TopBottomValue { get; set; }
    public XLTopBottomType TopBottomType { get; set; }
    public XLTopBottomPart TopBottomPart { get; set; }

    public XLFilterDynamicType DynamicType { get; set; }

    /// <summary>
    /// Basically average for dynamic filters. Value is refreshed during filter reapply.
    /// </summary>
    public Double DynamicValue { get; set; } = double.NaN;

    #endregion IXLFilterColumn Members

    /// <summary>
    /// A filter value used by top/bottom filter to compare with cell value.
    /// </summary>
    internal double TopBottomFilterValue { get; private set; } = double.NaN;

    public IEnumerator<XLFilter> GetEnumerator() => this._filters.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private void SetTopBottom(
        Int32 percentOrItemCount,
        XLTopBottomType type,
        Boolean takeTop,
        Boolean reapply
    )
    {
        if (percentOrItemCount is < 1 or > 500)
        {
            throw new ArgumentOutOfRangeException(
                nameof(percentOrItemCount),
                "Value must be between 1 and 500."
            );
        }

        this.ResetFilter(XLFilterType.TopBottom);
        this.TopBottomValue = percentOrItemCount;
        this.TopBottomType = type;
        this.TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;

        this.AddFilter(XLFilter.CreateTopBottom(takeTop, percentOrItemCount), reapply);
    }

    private double GetTopBottomFilterValue(XLTopBottomType type, int value, bool takeTop)
    {
        IXLRangeColumn column = this._autoFilter.Range.Column(this._column);
        IXLRangeColumn subColumn = column.Column(2, column.CellCount());
        IEnumerable<double> columnNumbers = subColumn
            .CellsUsed(c => c.CachedValue.IsUnifiedNumber)
            .Select(c => c.CachedValue.GetUnifiedNumber());
        Comparer<double> comparer = takeTop
            ? Comparer<double>.Create((x, y) => -x.CompareTo(y))
            : Comparer<double>.Create((x, y) => x.CompareTo(y));

        switch (type)
        {
            case XLTopBottomType.Items:
                int itemCount = value;
                return columnNumbers
                    .OrderBy(d => d, comparer)
                    .Take(itemCount)
                    .DefaultIfEmpty(double.NaN)
                    .LastOrDefault();
            case XLTopBottomType.Percent:
                int percent = value;
                double[] materializedNumbers = [.. columnNumbers];

                // Ceiling, so there is always at least one item.
                int itemCountByPercents = (int)
                    Math.Ceiling(materializedNumbers.Length * (double)percent / 100);
                return materializedNumbers
                    .OrderBy(d => d, comparer)
                    .Take(itemCountByPercents)
                    .DefaultIfEmpty(Double.NaN)
                    .LastOrDefault();
            default:
                throw new NotSupportedException();
        }
    }

    private void SetAverage(Boolean aboveAverage, Boolean reapply)
    {
        this.ResetFilter(XLFilterType.Dynamic);
        this.DynamicType = aboveAverage
            ? XLFilterDynamicType.AboveAverage
            : XLFilterDynamicType.BelowAverage;

        // `Average` is recalculated during reapply, so no need to calculate it twice.
        this.DynamicValue = reapply ? double.NaN : this.GetAverageFilterValue();
        this.AddFilter(XLFilter.CreateAverage(this.DynamicValue, aboveAverage), reapply);
    }

    private double GetAverageFilterValue()
    {
        IXLRangeColumn column = this._autoFilter.Range.Column(this._column);
        IXLRangeColumn subColumn = column.Column(2, column.CellCount());
        return subColumn
            .CellsUsed(c => c.CachedValue.IsUnifiedNumber)
            .Select(c => c.CachedValue.GetUnifiedNumber())
            .DefaultIfEmpty(Double.NaN)
            .Average();
    }

    private IXLFilterConnector AddCustomFilter(
        XLCellValue value,
        XLFilterOperator op,
        Boolean reapply
    )
    {
        this.ResetFilter(XLFilterType.Custom);
        this.AddFilter(XLFilter.CreateCustomFilter(value, op, XLConnector.Or), reapply);
        return new XLFilterConnector(this);
    }

    private IXLFilterConnector AddCustomFilter(string pattern, bool match, bool reapply)
    {
        this.ResetFilter(XLFilterType.Custom);
        this.AddFilter(XLFilter.CreateCustomPatternFilter(pattern, match, XLConnector.Or), reapply);
        return new XLFilterConnector(this);
    }

    private void ResetFilter(XLFilterType type)
    {
        this.Clear(false);
        this._autoFilter.IsEnabled = true;
        this.FilterType = type;
    }

    private void SwitchFilter(XLFilterType type)
    {
        this._autoFilter.IsEnabled = true;
        if (this.FilterType == type)
        {
            return;
        }

        this.Clear(false);
        this.FilterType = type;
    }

    internal void AddFilter(XLFilter filter, bool reapply = false)
    {
        int maxFilters = this.FilterType switch
        {
            XLFilterType.None => 0,
            XLFilterType.Regular => int.MaxValue,
            XLFilterType.Custom => 2,
            XLFilterType.TopBottom => 1,
            XLFilterType.Dynamic => 1,
            _ => throw new NotSupportedException(),
        };
        if (this._filters.Count >= maxFilters)
        {
            throw new InvalidOperationException(
                $"{this.FilterType} filter can have max {maxFilters} conditions."
            );
        }

        this._filters.Add(filter);
        if (reapply)
        {
            this._autoFilter.Reapply();
        }
    }

    internal void Refresh()
    {
        if (this.FilterType == XLFilterType.Dynamic)
        {
            // Update average value of a filter, so it is saved correctly and filter uses
            // correct value, even is cell values changed and avg was stale.
            this.DynamicValue = this.GetAverageFilterValue();
            this._filters[0].Value = this.DynamicValue;
        }

        if (this.FilterType == XLFilterType.TopBottom)
        {
            bool takeTop = this.TopBottomPart == XLTopBottomPart.Top;
            this.TopBottomFilterValue = this.GetTopBottomFilterValue(
                this.TopBottomType,
                this.TopBottomValue,
                takeTop
            );
        }
    }

    internal bool Check(IXLCell cell)
    {
        if (this._filters.Count == 0)
        {
            return true;
        }

        if (this._filters.Count == 1)
        {
            return this._filters[0].Condition(cell, this);
        }

        // All filter conditions are connected by a single type of logical condition. Regular
        // filters use 'Or', custom has up to two clauses connected by 'And'/'Or' and rest is
        // single clause.
        XLConnector connector = this._filters[1].Connector;
        return connector switch
        {
            XLConnector.And => this._filters.All(filter => filter.Condition(cell, this)),
            XLConnector.Or => this._filters.Any(filter => filter.Condition(cell, this)),
            _ => throw new NotSupportedException(),
        };
    }
}

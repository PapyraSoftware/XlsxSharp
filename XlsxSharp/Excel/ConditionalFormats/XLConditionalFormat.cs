#nullable disable warnings

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Misc;

namespace XlsxSharp.Excel.ConditionalFormats;

internal class XLConditionalFormat : IXLDxfContainer, IXLConditionalFormat
{
    private readonly XLWorksheet _worksheet;

    private sealed class NoRangeCfComparer : IEqualityComparer<XLConditionalFormat>
    {
        public bool Equals(XLConditionalFormat? xx, XLConditionalFormat? yy)
        {
            if (ReferenceEquals(xx, yy))
            {
                return true;
            }

            if (ReferenceEquals(xx, null))
            {
                return false;
            }

            if (ReferenceEquals(yy, null))
            {
                return false;
            }

            if (xx.GetType() != yy.GetType())
            {
                return false;
            }

            IEnumerable<string>? xxValues = xx
                .Values.Values.Where(v => v == null || !v.IsFormula)
                .Select(v => v?.Value);
            IEnumerable<string>? yyValues = yy
                .Values.Values.Where(v => v == null || !v.IsFormula)
                .Select(v => v?.Value);
            IEnumerable<string>? xxFormulas =
                xx.Areas.Count > 0
                    ? xx
                        .Values.Values.Where(v => v != null && v.IsFormula)
                        .Select(f => ((XLCell)xx.Range.FirstCell()).GetFormulaR1C1(f.Value))
                    : null;
            IEnumerable<string>? yyFormulas =
                yy.Areas.Count > 0
                    ? yy
                        .Values.Values.Where(v => v != null && v.IsFormula)
                        .Select(f => ((XLCell)yy.Range.FirstCell()).GetFormulaR1C1(f.Value))
                    : null;
            XLDxfValue? xStyle = xx.FormatValue;
            XLDxfValue? yStyle = yy.FormatValue;
            return Equals(xStyle, yStyle)
                && xx.ConditionalFormatType == yy.ConditionalFormatType
                && xx.TimePeriod == yy.TimePeriod
                && xx.IconSetStyle == yy.IconSetStyle
                && xx.Operator == yy.Operator
                && xx.Bottom == yy.Bottom
                && xx.Percent == yy.Percent
                && xx.ReverseIconOrder == yy.ReverseIconOrder
                && xx.StopIfTrue == yy.StopIfTrue
                && xx.ShowIconOnly == yy.ShowIconOnly
                && xx.ShowBarOnly == yy.ShowBarOnly
                && SetEquals(xxValues, yyValues)
                && SetEquals(xxFormulas, yyFormulas)
                && Equals(xx.Colors, yy.Colors)
                && Equals(xx.ContentTypes, yy.ContentTypes)
                && Equals(xx.IconSetOperators, yy.IconSetOperators);
        }

        public int GetHashCode(XLConditionalFormat obj)
        {
            XLConditionalFormat? xx = obj;
            XLDxfValue? xStyle = obj.FormatValue;
            IEnumerable<string>? xValues = xx
                .Values.Values.Where(v => !v.IsFormula)
                .Select(v => v.Value);
            if (obj.Areas.Count > 0)
            {
                xValues = xValues.Union(
                    xx.Values.Values.Where(v => v.IsFormula)
                        .Select(f => ((XLCell)obj.Range.FirstCell()).GetFormulaR1C1(f.Value))
                );
            }
            unchecked
            {
                int hashCode = xStyle.GetHashCode();
                hashCode = (hashCode * 397) ^ xValues.GetHashCode();
                hashCode = (hashCode * 397) ^ (xx.Colors != null ? xx.Colors.GetHashCode() : 0);
                hashCode =
                    (hashCode * 397)
                    ^ (xx.ContentTypes != null ? xx.ContentTypes.GetHashCode() : 0);
                hashCode =
                    (hashCode * 397)
                    ^ (xx.IconSetOperators != null ? xx.IconSetOperators.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (int)xx.ConditionalFormatType;
                hashCode = (hashCode * 397) ^ (int)xx.TimePeriod;
                hashCode = (hashCode * 397) ^ (int)xx.IconSetStyle;
                hashCode = (hashCode * 397) ^ (int)xx.Operator;
                hashCode = (hashCode * 397) ^ xx.Bottom.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.Percent.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ReverseIconOrder.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowIconOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.ShowBarOnly.GetHashCode();
                hashCode = (hashCode * 397) ^ xx.StopIfTrue.GetHashCode();
                return hashCode;
            }
        }

        private static bool SetEquals<T>(IEnumerable<T> first, IEnumerable<T> second) =>
            new HashSet<T>(second, EqualityComparer<T>.Default).SetEquals(first);

        private static bool Equals<TValue>(Dictionary<int, TValue> x, Dictionary<int, TValue> y)
        {
            if (x.Count != y.Count)
            {
                return false;
            }

            if (x.Keys.Except(y.Keys).Any())
            {
                return false;
            }

            if (y.Keys.Except(x.Keys).Any())
            {
                return false;
            }

            EqualityComparer<TValue>? valueComparer = EqualityComparer<TValue>.Default;
            foreach (KeyValuePair<int, TValue> pair in x)
            {
                if (!valueComparer.Equals(pair.Value, y[pair.Key]))
                {
                    return false;
                }
            }

            return true;
        }
    }

    internal static IEqualityComparer<XLConditionalFormat> NoRangeComparer { get; } =
        new NoRangeCfComparer();

    #region Constructors

    internal XLConditionalFormat(XLWorksheet worksheet, XLAreaList areaList)
    {
        this._worksheet = worksheet;
        this.Id = Guid.NewGuid();
        this.Areas = areaList;
        this.Values = new XLDictionary<XLFormula>();
        this.Colors = new XLDictionary<XLColor>();
        this.ContentTypes = new XLDictionary<XLCFContentType>();
        this.IconSetOperators = new XLDictionary<XLCFIconSetOperator>();
    }

    /// <summary>
    /// Copy ctor.
    /// </summary>
    internal XLConditionalFormat(
        XLWorksheet worksheet,
        XLConditionalFormat other,
        XLAreaList areaList
    )
        : this(worksheet, areaList)
    {
        XLDxfValue? otherDxf = other.FormatValue;
        this.FormatValue = otherDxf is not null
            ? this._worksheet.Workbook.Styles.GetRegisteredDxFormat(otherDxf, static x => x)
            : null;
        this.ConditionalFormatType = other.ConditionalFormatType;
        this.TimePeriod = other.TimePeriod;
        this.IconSetStyle = other.IconSetStyle;
        this.Operator = other.Operator;
        this.Bottom = other.Bottom;
        this.Percent = other.Percent;
        this.ReverseIconOrder = other.ReverseIconOrder;
        this.ShowIconOnly = other.ShowIconOnly;
        this.ShowBarOnly = other.ShowBarOnly;
        this.StopIfTrue = other.StopIfTrue;

        Point sourceAnchor = other.Areas[0].FirstPoint;
        Point targetAnchor = this.Areas[0].FirstPoint;
        foreach ((int key, XLFormula? originalValue) in other.Values)
        {
            this.Values.Add(key, originalValue.GetAdjustedCopy(sourceAnchor, targetAnchor));
        }

        this.Colors = other.Colors.CopyDictionary();
        this.ContentTypes = other.ContentTypes.CopyDictionary();
        this.IconSetOperators = other.IconSetOperators.CopyDictionary();
    }

    #endregion Constructors

    public Guid Id { get; internal set; }

    /// <summary>
    /// Priority of formatting rule. Lower values have higher priority than higher values.
    /// Minimum value is 1. It is basically used for ordering of CF during saving.
    /// </summary>
    internal int Priority { get; set; }

    public XLDxfValue? FormatValue { get; set; }

    internal XLDxFormat Format => new(this._worksheet.Workbook.Styles, this);

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    public XLDictionary<XLFormula> Values { get; }

    public XLDictionary<XLColor> Colors { get; }

    public XLDictionary<XLCFContentType> ContentTypes { get; }

    public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; }

    public IXLRange Range
    {
        get => this._worksheet.Range(this.Areas[0]);
        set => this.Areas = XLAreaList.FromRange(this._worksheet, value);
    }

    public IEnumerable<IXLRange> Ranges
    {
        get
        {
            XLRanges? ranges = new(this._worksheet);
            foreach (Area area in this.Areas)
            {
                ranges.Add(this._worksheet.Range(area));
            }

            return ranges;
        }
        set => this.Areas = XLAreaList.FromRanges(this._worksheet, value);
    }

    public XLConditionalFormatType ConditionalFormatType { get; set; }

    public XLTimePeriod TimePeriod { get; set; }

    public XLIconSetStyle IconSetStyle { get; set; }

    public XLCFOperator Operator { get; set; }

    public bool Bottom { get; set; }

    public bool Percent { get; set; }

    public bool ReverseIconOrder { get; set; }

    public bool ShowIconOnly { get; set; }

    public bool ShowBarOnly { get; set; }

    public bool StopIfTrue { get; set; }

    internal XLAreaList Areas { get; set; }

    public IXLConditionalFormat SetStopIfTrue() => this.SetStopIfTrue(true);

    public IXLConditionalFormat SetStopIfTrue(bool value)
    {
        this.StopIfTrue = value;
        return this;
    }

    public IXLConditionalFormat CopyTo(IXLWorksheet targetSheet)
    {
        if (targetSheet == this.Range?.Worksheet)
        {
            throw new InvalidOperationException(
                "Cannot copy conditional format to the worksheet it already belongs to."
            );
        }

        XLConditionalFormat? newCf = new((XLWorksheet)targetSheet, this, this.Areas);
        targetSheet.ConditionalFormats.Add(newCf);
        return newCf;
    }

    public IXLStyle WhenIsBlank()
    {
        this.ConditionalFormatType = XLConditionalFormatType.IsBlank;
        return this.Style;
    }

    public IXLStyle WhenNotBlank()
    {
        this.ConditionalFormatType = XLConditionalFormatType.NotBlank;
        return this.Style;
    }

    public IXLStyle WhenIsError()
    {
        this.ConditionalFormatType = XLConditionalFormatType.IsError;
        return this.Style;
    }

    public IXLStyle WhenNotError()
    {
        this.ConditionalFormatType = XLConditionalFormatType.NotError;
        return this.Style;
    }

    public IXLStyle WhenDateIs(XLTimePeriod timePeriod)
    {
        this.TimePeriod = timePeriod;
        this.ConditionalFormatType = XLConditionalFormatType.TimePeriod;
        return this.Style;
    }

    public IXLStyle WhenContains(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.ConditionalFormatType = XLConditionalFormatType.ContainsText;
        this.Operator = XLCFOperator.Contains;
        return this.Style;
    }

    public IXLStyle WhenNotContains(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.ConditionalFormatType = XLConditionalFormatType.NotContainsText;
        this.Operator = XLCFOperator.NotContains;
        return this.Style;
    }

    public IXLStyle WhenStartsWith(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.ConditionalFormatType = XLConditionalFormatType.StartsWith;
        this.Operator = XLCFOperator.StartsWith;
        return this.Style;
    }

    public IXLStyle WhenEndsWith(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.ConditionalFormatType = XLConditionalFormatType.EndsWith;
        this.Operator = XLCFOperator.EndsWith;
        return this.Style;
    }

    public IXLStyle WhenEquals(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.Equal;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenNotEquals(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.NotEqual;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenGreaterThan(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.GreaterThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenLessThan(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.LessThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenEqualOrGreaterThan(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.EqualOrGreaterThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenEqualOrLessThan(string value)
    {
        this.Values.Initialize(new XLFormula { Value = value });
        this.Operator = XLCFOperator.EqualOrLessThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenBetween(string minValue, string maxValue)
    {
        this.Values.Initialize(new XLFormula { Value = minValue });
        this.Values.Add(new XLFormula { Value = maxValue });
        this.Operator = XLCFOperator.Between;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenNotBetween(string minValue, string maxValue)
    {
        this.Values.Initialize(new XLFormula { Value = minValue });
        this.Values.Add(new XLFormula { Value = maxValue });
        this.Operator = XLCFOperator.NotBetween;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenEquals(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.Equal;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenNotEquals(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.NotEqual;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenGreaterThan(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.GreaterThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenLessThan(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.LessThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenEqualOrGreaterThan(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.EqualOrGreaterThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenEqualOrLessThan(double value)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Operator = XLCFOperator.EqualOrLessThan;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenBetween(double minValue, double maxValue)
    {
        this.Values.Initialize(new XLFormula(minValue));
        this.Values.Add(new XLFormula(maxValue));
        this.Operator = XLCFOperator.Between;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenNotBetween(double minValue, double maxValue)
    {
        this.Values.Initialize(new XLFormula(minValue));
        this.Values.Add(new XLFormula(maxValue));
        this.Operator = XLCFOperator.NotBetween;
        this.ConditionalFormatType = XLConditionalFormatType.CellIs;
        return this.Style;
    }

    public IXLStyle WhenIsDuplicate()
    {
        this.ConditionalFormatType = XLConditionalFormatType.IsDuplicate;
        return this.Style;
    }

    public IXLStyle WhenIsUnique()
    {
        this.ConditionalFormatType = XLConditionalFormatType.IsUnique;
        return this.Style;
    }

    public IXLStyle WhenIsTrue(string formula)
    {
        string f = formula.TrimStart()[0] == '=' ? formula : "=" + formula;
        this.Values.Initialize(new XLFormula { Value = f });
        this.ConditionalFormatType = XLConditionalFormatType.Expression;
        return this.Style;
    }

    public IXLStyle WhenIsTop(int value, XLTopBottomType topBottomType = XLTopBottomType.Items)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Percent = topBottomType == XLTopBottomType.Percent;
        this.ConditionalFormatType = XLConditionalFormatType.Top10;
        this.Bottom = false;
        return this.Style;
    }

    public IXLStyle WhenIsBottom(int value, XLTopBottomType topBottomType = XLTopBottomType.Items)
    {
        this.Values.Initialize(new XLFormula(value));
        this.Percent = topBottomType == XLTopBottomType.Percent;
        this.ConditionalFormatType = XLConditionalFormatType.Top10;
        this.Bottom = true;
        return this.Style;
    }

    public IXLCFColorScaleMin ColorScale()
    {
        this.ConditionalFormatType = XLConditionalFormatType.ColorScale;
        return new XLCFColorScaleMin(this);
    }

    public IXLCFDataBarMin DataBar(XLColor color, bool showBarOnly = false)
    {
        this.Colors.Initialize(color);
        this.ShowBarOnly = showBarOnly;
        this.ConditionalFormatType = XLConditionalFormatType.DataBar;
        return new XLCFDataBarMin(this);
    }

    public IXLCFDataBarMin DataBar(
        XLColor positiveColor,
        XLColor negativeColor,
        bool showBarOnly = false
    )
    {
        this.Colors.Initialize(positiveColor);
        this.Colors.Add(negativeColor);
        this.ShowBarOnly = showBarOnly;
        this.ConditionalFormatType = XLConditionalFormatType.DataBar;
        return new XLCFDataBarMin(this);
    }

    public IXLCFIconSet IconSet(
        XLIconSetStyle iconSetStyle,
        bool reverseIconOrder = false,
        bool showIconOnly = false
    )
    {
        this.IconSetOperators.Clear();
        this.Values.Clear();
        this.ContentTypes.Clear();
        this.ConditionalFormatType = XLConditionalFormatType.IconSet;
        this.IconSetStyle = iconSetStyle;
        this.ReverseIconOrder = reverseIconOrder;
        this.ShowIconOnly = showIconOnly;
        return new XLCFIconSet(this);
    }
}

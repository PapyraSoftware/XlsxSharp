#nullable disable

using XlsxSharp.Excel.Misc;

namespace XlsxSharp.Excel.ConditionalFormats;

public enum XLTimePeriod
{
    Yesterday,
    Today,
    Tomorrow,
    InTheLast7Days,
    LastWeek,
    ThisWeek,
    NextWeek,
    LastMonth,
    ThisMonth,
    NextMonth,
}

public enum XLIconSetStyle
{
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights1,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}

public enum XLConditionalFormatType
{
    Expression,
    CellIs,
    ColorScale,
    DataBar,
    IconSet,
    Top10,
    IsUnique,
    IsDuplicate,
    ContainsText,
    NotContainsText,
    StartsWith,
    EndsWith,
    IsBlank,
    NotBlank,
    IsError,
    NotError,
    TimePeriod,
    AboveAverage,
}

public enum XLCFOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    LessThan,
    EqualOrGreaterThan,
    EqualOrLessThan,
    Between,
    NotBetween,
    Contains,
    NotContains,
    StartsWith,
    EndsWith,
}

public interface IXLConditionalFormat
{
    public IXLStyle Style { get; set; }

    public IXLStyle WhenIsBlank();

    public IXLStyle WhenNotBlank();

    public IXLStyle WhenIsError();

    public IXLStyle WhenNotError();

    public IXLStyle WhenDateIs(XLTimePeriod timePeriod);

    public IXLStyle WhenContains(string value);

    public IXLStyle WhenNotContains(string value);

    public IXLStyle WhenStartsWith(string value);

    public IXLStyle WhenEndsWith(string value);

    public IXLStyle WhenEquals(string value);

    public IXLStyle WhenNotEquals(string value);

    public IXLStyle WhenGreaterThan(string value);

    public IXLStyle WhenLessThan(string value);

    public IXLStyle WhenEqualOrGreaterThan(string value);

    public IXLStyle WhenEqualOrLessThan(string value);

    public IXLStyle WhenBetween(string minValue, string maxValue);

    public IXLStyle WhenNotBetween(string minValue, string maxValue);

    public IXLStyle WhenEquals(double value);

    public IXLStyle WhenNotEquals(double value);

    public IXLStyle WhenGreaterThan(double value);

    public IXLStyle WhenLessThan(double value);

    public IXLStyle WhenEqualOrGreaterThan(double value);

    public IXLStyle WhenEqualOrLessThan(double value);

    public IXLStyle WhenBetween(double minValue, double maxValue);

    public IXLStyle WhenNotBetween(double minValue, double maxValue);

    public IXLStyle WhenIsDuplicate();

    public IXLStyle WhenIsUnique();

    public IXLStyle WhenIsTrue(string formula);

    public IXLStyle WhenIsTop(int value, XLTopBottomType topBottomType = XLTopBottomType.Items);

    public IXLStyle WhenIsBottom(int value, XLTopBottomType topBottomType);

    public IXLCFColorScaleMin ColorScale();

    public IXLCFDataBarMin DataBar(XLColor color, bool showBarOnly = false);

    public IXLCFDataBarMin DataBar(
        XLColor positiveColor,
        XLColor negativeColor,
        bool showBarOnly = false
    );

    public IXLCFIconSet IconSet(
        XLIconSetStyle iconSetStyle,
        bool reverseIconOrder = false,
        bool showIconOnly = false
    );

    public XLConditionalFormatType ConditionalFormatType { get; }

    public XLIconSetStyle IconSetStyle { get; }

    public XLTimePeriod TimePeriod { get; }

    public bool ReverseIconOrder { get; }

    public bool ShowIconOnly { get; }

    public bool ShowBarOnly { get; }

    public bool StopIfTrue { get; }

    /// <summary>
    /// The first of the <see cref="Ranges"/>.
    /// </summary>
    public IXLRange Range { get; set; }

    /// <summary>
    /// Get or set all ranges the conditional format applies to.
    /// </summary>
    /// <exception cref="ArgumentException">If sequence contains no elements or the range is from different worksheet.</exception>
    public IEnumerable<IXLRange> Ranges { get; set; }

    public XLDictionary<XLFormula> Values { get; }

    public XLDictionary<XLColor> Colors { get; }

    public XLDictionary<XLCFContentType> ContentTypes { get; }

    public XLDictionary<XLCFIconSetOperator> IconSetOperators { get; }

    public XLCFOperator Operator { get; }

    public bool Bottom { get; }

    public bool Percent { get; }

    public IXLConditionalFormat SetStopIfTrue();

    public IXLConditionalFormat SetStopIfTrue(bool value);

    public IXLConditionalFormat CopyTo(IXLWorksheet targetSheet);
}

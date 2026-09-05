using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Tables;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// The enumerations of <c>xl/worksheets/sheetN.xml</c> and the parts that hang off it (tables,
/// drawings and the x14 sparkline extension), keyed by the string OOXML writes for them.
/// </summary>
/// <remarks>
/// These replace the SDK enum values the worksheet reader and writer converted through. The
/// strings are the ones the SDK serialises, and the tests check every entry in both directions
/// against the conversion it replaces, so a wrong spelling shows up as a failing test rather than
/// as a worksheet that silently loads with the wrong page order or the wrong icon set.
/// </remarks>
internal static class WorksheetXmlEnums
{
    #region Page setup

    internal static XLPageOrientation ParsePageOrientation(string value) =>
        value switch
        {
            "default" => XLPageOrientation.Default,
            "portrait" => XLPageOrientation.Portrait,
            "landscape" => XLPageOrientation.Landscape,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLPageOrientation value) =>
        value switch
        {
            XLPageOrientation.Default => "default",
            XLPageOrientation.Portrait => "portrait",
            XLPageOrientation.Landscape => "landscape",
            _ => throw UnknownValue(value),
        };

    internal static XLPageOrderValues ParsePageOrder(string value) =>
        value switch
        {
            "downThenOver" => XLPageOrderValues.DownThenOver,
            "overThenDown" => XLPageOrderValues.OverThenDown,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLPageOrderValues value) =>
        value switch
        {
            XLPageOrderValues.DownThenOver => "downThenOver",
            XLPageOrderValues.OverThenDown => "overThenDown",
            _ => throw UnknownValue(value),
        };

    internal static XLShowCommentsValues ParseShowComments(string value) =>
        value switch
        {
            "none" => XLShowCommentsValues.None,
            "asDisplayed" => XLShowCommentsValues.AsDisplayed,
            "atEnd" => XLShowCommentsValues.AtEnd,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLShowCommentsValues value) =>
        value switch
        {
            XLShowCommentsValues.None => "none",
            XLShowCommentsValues.AsDisplayed => "asDisplayed",
            XLShowCommentsValues.AtEnd => "atEnd",
            _ => throw UnknownValue(value),
        };

    internal static XLPrintErrorValues ParsePrintError(string value) =>
        value switch
        {
            "displayed" => XLPrintErrorValues.Displayed,
            "blank" => XLPrintErrorValues.Blank,
            "dash" => XLPrintErrorValues.Dash,
            "NA" => XLPrintErrorValues.NA,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLPrintErrorValues value) =>
        value switch
        {
            XLPrintErrorValues.Displayed => "displayed",
            XLPrintErrorValues.Blank => "blank",
            XLPrintErrorValues.Dash => "dash",
            XLPrintErrorValues.NA => "NA",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Sheet views

    internal static XLSheetViewOptions ParseSheetViewType(string value) =>
        value switch
        {
            "normal" => XLSheetViewOptions.Normal,
            "pageBreakPreview" => XLSheetViewOptions.PageBreakPreview,
            "pageLayout" => XLSheetViewOptions.PageLayout,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLSheetViewOptions value) =>
        value switch
        {
            XLSheetViewOptions.Normal => "normal",
            XLSheetViewOptions.PageBreakPreview => "pageBreakPreview",
            XLSheetViewOptions.PageLayout => "pageLayout",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Data validation

    internal static XLAllowedValues ParseAllowedValues(string value) =>
        value switch
        {
            "none" => XLAllowedValues.AnyValue,
            "whole" => XLAllowedValues.WholeNumber,
            "decimal" => XLAllowedValues.Decimal,
            "list" => XLAllowedValues.List,
            "date" => XLAllowedValues.Date,
            "time" => XLAllowedValues.Time,
            "textLength" => XLAllowedValues.TextLength,
            "custom" => XLAllowedValues.Custom,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLAllowedValues value) =>
        value switch
        {
            XLAllowedValues.AnyValue => "none",
            XLAllowedValues.WholeNumber => "whole",
            XLAllowedValues.Decimal => "decimal",
            XLAllowedValues.List => "list",
            XLAllowedValues.Date => "date",
            XLAllowedValues.Time => "time",
            XLAllowedValues.TextLength => "textLength",
            XLAllowedValues.Custom => "custom",
            _ => throw UnknownValue(value),
        };

    internal static XLErrorStyle ParseErrorStyle(string value) =>
        value switch
        {
            "stop" => XLErrorStyle.Stop,
            "warning" => XLErrorStyle.Warning,
            "information" => XLErrorStyle.Information,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLErrorStyle value) =>
        value switch
        {
            XLErrorStyle.Stop => "stop",
            XLErrorStyle.Warning => "warning",
            XLErrorStyle.Information => "information",
            _ => throw UnknownValue(value),
        };

    internal static XLOperator ParseDataValidationOperator(string value) =>
        value switch
        {
            "between" => XLOperator.Between,
            "notBetween" => XLOperator.NotBetween,
            "equal" => XLOperator.EqualTo,
            "notEqual" => XLOperator.NotEqualTo,
            "lessThan" => XLOperator.LessThan,
            "lessThanOrEqual" => XLOperator.EqualOrLessThan,
            "greaterThan" => XLOperator.GreaterThan,
            "greaterThanOrEqual" => XLOperator.EqualOrGreaterThan,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLOperator value) =>
        value switch
        {
            XLOperator.Between => "between",
            XLOperator.NotBetween => "notBetween",
            XLOperator.EqualTo => "equal",
            XLOperator.NotEqualTo => "notEqual",
            XLOperator.LessThan => "lessThan",
            XLOperator.EqualOrLessThan => "lessThanOrEqual",
            XLOperator.GreaterThan => "greaterThan",
            XLOperator.EqualOrGreaterThan => "greaterThanOrEqual",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Conditional formats

    internal static XLConditionalFormatType ParseConditionalFormatType(string value) =>
        value switch
        {
            "expression" => XLConditionalFormatType.Expression,
            "cellIs" => XLConditionalFormatType.CellIs,
            "colorScale" => XLConditionalFormatType.ColorScale,
            "dataBar" => XLConditionalFormatType.DataBar,
            "iconSet" => XLConditionalFormatType.IconSet,
            "top10" => XLConditionalFormatType.Top10,
            "uniqueValues" => XLConditionalFormatType.IsUnique,
            "duplicateValues" => XLConditionalFormatType.IsDuplicate,
            "containsText" => XLConditionalFormatType.ContainsText,
            "notContainsText" => XLConditionalFormatType.NotContainsText,
            "beginsWith" => XLConditionalFormatType.StartsWith,
            "endsWith" => XLConditionalFormatType.EndsWith,
            "containsBlanks" => XLConditionalFormatType.IsBlank,
            "notContainsBlanks" => XLConditionalFormatType.NotBlank,
            "containsErrors" => XLConditionalFormatType.IsError,
            "notContainsErrors" => XLConditionalFormatType.NotError,
            "timePeriod" => XLConditionalFormatType.TimePeriod,
            "aboveAverage" => XLConditionalFormatType.AboveAverage,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLConditionalFormatType value) =>
        value switch
        {
            XLConditionalFormatType.Expression => "expression",
            XLConditionalFormatType.CellIs => "cellIs",
            XLConditionalFormatType.ColorScale => "colorScale",
            XLConditionalFormatType.DataBar => "dataBar",
            XLConditionalFormatType.IconSet => "iconSet",
            XLConditionalFormatType.Top10 => "top10",
            XLConditionalFormatType.IsUnique => "uniqueValues",
            XLConditionalFormatType.IsDuplicate => "duplicateValues",
            XLConditionalFormatType.ContainsText => "containsText",
            XLConditionalFormatType.NotContainsText => "notContainsText",
            XLConditionalFormatType.StartsWith => "beginsWith",
            XLConditionalFormatType.EndsWith => "endsWith",
            XLConditionalFormatType.IsBlank => "containsBlanks",
            XLConditionalFormatType.NotBlank => "notContainsBlanks",
            XLConditionalFormatType.IsError => "containsErrors",
            XLConditionalFormatType.NotError => "notContainsErrors",
            XLConditionalFormatType.TimePeriod => "timePeriod",
            XLConditionalFormatType.AboveAverage => "aboveAverage",
            _ => throw UnknownValue(value),
        };

    internal static XLCFOperator ParseCfOperator(string value) =>
        value switch
        {
            "lessThan" => XLCFOperator.LessThan,
            "lessThanOrEqual" => XLCFOperator.EqualOrLessThan,
            "equal" => XLCFOperator.Equal,
            "notEqual" => XLCFOperator.NotEqual,
            "greaterThanOrEqual" => XLCFOperator.EqualOrGreaterThan,
            "greaterThan" => XLCFOperator.GreaterThan,
            "between" => XLCFOperator.Between,
            "notBetween" => XLCFOperator.NotBetween,
            "containsText" => XLCFOperator.Contains,
            "notContains" => XLCFOperator.NotContains,
            "beginsWith" => XLCFOperator.StartsWith,
            "endsWith" => XLCFOperator.EndsWith,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLCFOperator value) =>
        value switch
        {
            XLCFOperator.LessThan => "lessThan",
            XLCFOperator.EqualOrLessThan => "lessThanOrEqual",
            XLCFOperator.Equal => "equal",
            XLCFOperator.NotEqual => "notEqual",
            XLCFOperator.EqualOrGreaterThan => "greaterThanOrEqual",
            XLCFOperator.GreaterThan => "greaterThan",
            XLCFOperator.Between => "between",
            XLCFOperator.NotBetween => "notBetween",
            XLCFOperator.Contains => "containsText",
            XLCFOperator.NotContains => "notContains",
            XLCFOperator.StartsWith => "beginsWith",
            XLCFOperator.EndsWith => "endsWith",
            _ => throw UnknownValue(value),
        };

    internal static XLCFContentType ParseCfContentType(string value) =>
        value switch
        {
            "num" => XLCFContentType.Number,
            "percent" => XLCFContentType.Percent,
            "max" => XLCFContentType.Maximum,
            "min" => XLCFContentType.Minimum,
            "formula" => XLCFContentType.Formula,
            "percentile" => XLCFContentType.Percentile,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLCFContentType value) =>
        value switch
        {
            XLCFContentType.Number => "num",
            XLCFContentType.Percent => "percent",
            XLCFContentType.Maximum => "max",
            XLCFContentType.Minimum => "min",
            XLCFContentType.Formula => "formula",
            XLCFContentType.Percentile => "percentile",
            _ => throw UnknownValue(value),
        };

    internal static XLIconSetStyle ParseIconSetStyle(string value) =>
        value switch
        {
            "3Arrows" => XLIconSetStyle.ThreeArrows,
            "3ArrowsGray" => XLIconSetStyle.ThreeArrowsGray,
            "3Flags" => XLIconSetStyle.ThreeFlags,
            "3TrafficLights1" => XLIconSetStyle.ThreeTrafficLights1,
            "3TrafficLights2" => XLIconSetStyle.ThreeTrafficLights2,
            "3Signs" => XLIconSetStyle.ThreeSigns,
            "3Symbols" => XLIconSetStyle.ThreeSymbols,
            "3Symbols2" => XLIconSetStyle.ThreeSymbols2,
            "4Arrows" => XLIconSetStyle.FourArrows,
            "4ArrowsGray" => XLIconSetStyle.FourArrowsGray,
            "4RedToBlack" => XLIconSetStyle.FourRedToBlack,
            "4Rating" => XLIconSetStyle.FourRating,
            "4TrafficLights" => XLIconSetStyle.FourTrafficLights,
            "5Arrows" => XLIconSetStyle.FiveArrows,
            "5ArrowsGray" => XLIconSetStyle.FiveArrowsGray,
            "5Rating" => XLIconSetStyle.FiveRating,
            "5Quarters" => XLIconSetStyle.FiveQuarters,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLIconSetStyle value) =>
        value switch
        {
            XLIconSetStyle.ThreeArrows => "3Arrows",
            XLIconSetStyle.ThreeArrowsGray => "3ArrowsGray",
            XLIconSetStyle.ThreeFlags => "3Flags",
            XLIconSetStyle.ThreeTrafficLights1 => "3TrafficLights1",
            XLIconSetStyle.ThreeTrafficLights2 => "3TrafficLights2",
            XLIconSetStyle.ThreeSigns => "3Signs",
            XLIconSetStyle.ThreeSymbols => "3Symbols",
            XLIconSetStyle.ThreeSymbols2 => "3Symbols2",
            XLIconSetStyle.FourArrows => "4Arrows",
            XLIconSetStyle.FourArrowsGray => "4ArrowsGray",
            XLIconSetStyle.FourRedToBlack => "4RedToBlack",
            XLIconSetStyle.FourRating => "4Rating",
            XLIconSetStyle.FourTrafficLights => "4TrafficLights",
            XLIconSetStyle.FiveArrows => "5Arrows",
            XLIconSetStyle.FiveArrowsGray => "5ArrowsGray",
            XLIconSetStyle.FiveRating => "5Rating",
            XLIconSetStyle.FiveQuarters => "5Quarters",
            _ => throw UnknownValue(value),
        };

    internal static XLTimePeriod ParseTimePeriod(string value) =>
        value switch
        {
            "today" => XLTimePeriod.Today,
            "yesterday" => XLTimePeriod.Yesterday,
            "tomorrow" => XLTimePeriod.Tomorrow,
            "last7Days" => XLTimePeriod.InTheLast7Days,
            "thisMonth" => XLTimePeriod.ThisMonth,
            "lastMonth" => XLTimePeriod.LastMonth,
            "nextMonth" => XLTimePeriod.NextMonth,
            "thisWeek" => XLTimePeriod.ThisWeek,
            "lastWeek" => XLTimePeriod.LastWeek,
            "nextWeek" => XLTimePeriod.NextWeek,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLTimePeriod value) =>
        value switch
        {
            XLTimePeriod.Today => "today",
            XLTimePeriod.Yesterday => "yesterday",
            XLTimePeriod.Tomorrow => "tomorrow",
            XLTimePeriod.InTheLast7Days => "last7Days",
            XLTimePeriod.ThisMonth => "thisMonth",
            XLTimePeriod.LastMonth => "lastMonth",
            XLTimePeriod.NextMonth => "nextMonth",
            XLTimePeriod.ThisWeek => "thisWeek",
            XLTimePeriod.LastWeek => "lastWeek",
            XLTimePeriod.NextWeek => "nextWeek",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Auto filter

    internal static XLFilterOperator ParseFilterOperator(string value) =>
        value switch
        {
            "equal" => XLFilterOperator.Equal,
            "notEqual" => XLFilterOperator.NotEqual,
            "lessThan" => XLFilterOperator.LessThan,
            "lessThanOrEqual" => XLFilterOperator.EqualOrLessThan,
            "greaterThan" => XLFilterOperator.GreaterThan,
            "greaterThanOrEqual" => XLFilterOperator.EqualOrGreaterThan,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLFilterOperator value) =>
        value switch
        {
            XLFilterOperator.Equal => "equal",
            XLFilterOperator.NotEqual => "notEqual",
            XLFilterOperator.LessThan => "lessThan",
            XLFilterOperator.EqualOrLessThan => "lessThanOrEqual",
            XLFilterOperator.GreaterThan => "greaterThan",
            XLFilterOperator.EqualOrGreaterThan => "greaterThanOrEqual",
            _ => throw UnknownValue(value),
        };

    /// <summary>
    /// The workbook model only knows the two average filters. The schema has three dozen more
    /// (the date periods and the Q1..Q4 and M1..M12 shorthands), and reading one of those is an
    /// error here just as it was before - the model has nowhere to put it.
    /// </summary>
    internal static XLFilterDynamicType ParseFilterDynamicType(string value) =>
        value switch
        {
            "aboveAverage" => XLFilterDynamicType.AboveAverage,
            "belowAverage" => XLFilterDynamicType.BelowAverage,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLFilterDynamicType value) =>
        value switch
        {
            XLFilterDynamicType.AboveAverage => "aboveAverage",
            XLFilterDynamicType.BelowAverage => "belowAverage",
            _ => throw UnknownValue(value),
        };

    internal static XLDateTimeGrouping ParseDateTimeGrouping(string value) =>
        value switch
        {
            "year" => XLDateTimeGrouping.Year,
            "month" => XLDateTimeGrouping.Month,
            "day" => XLDateTimeGrouping.Day,
            "hour" => XLDateTimeGrouping.Hour,
            "minute" => XLDateTimeGrouping.Minute,
            "second" => XLDateTimeGrouping.Second,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLDateTimeGrouping value) =>
        value switch
        {
            XLDateTimeGrouping.Year => "year",
            XLDateTimeGrouping.Month => "month",
            XLDateTimeGrouping.Day => "day",
            XLDateTimeGrouping.Hour => "hour",
            XLDateTimeGrouping.Minute => "minute",
            XLDateTimeGrouping.Second => "second",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Tables

    internal static XLTotalsRowFunction ParseTotalsRowFunction(string value) =>
        value switch
        {
            "none" => XLTotalsRowFunction.None,
            "sum" => XLTotalsRowFunction.Sum,
            "min" => XLTotalsRowFunction.Minimum,
            "max" => XLTotalsRowFunction.Maximum,
            "average" => XLTotalsRowFunction.Average,
            "count" => XLTotalsRowFunction.Count,
            "countNums" => XLTotalsRowFunction.CountNumbers,
            "stdDev" => XLTotalsRowFunction.StandardDeviation,
            "var" => XLTotalsRowFunction.Variance,
            "custom" => XLTotalsRowFunction.Custom,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLTotalsRowFunction value) =>
        value switch
        {
            XLTotalsRowFunction.None => "none",
            XLTotalsRowFunction.Sum => "sum",
            XLTotalsRowFunction.Minimum => "min",
            XLTotalsRowFunction.Maximum => "max",
            XLTotalsRowFunction.Average => "average",
            XLTotalsRowFunction.Count => "count",
            XLTotalsRowFunction.CountNumbers => "countNums",
            XLTotalsRowFunction.StandardDeviation => "stdDev",
            XLTotalsRowFunction.Variance => "var",
            XLTotalsRowFunction.Custom => "custom",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Phonetics

    internal static XLPhoneticType ParsePhoneticType(string value) =>
        value switch
        {
            "halfwidthKatakana" => XLPhoneticType.HalfWidthKatakana,
            "fullwidthKatakana" => XLPhoneticType.FullWidthKatakana,
            "Hiragana" => XLPhoneticType.Hiragana,
            "noConversion" => XLPhoneticType.NoConversion,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLPhoneticType value) =>
        value switch
        {
            XLPhoneticType.HalfWidthKatakana => "halfwidthKatakana",
            XLPhoneticType.FullWidthKatakana => "fullwidthKatakana",
            XLPhoneticType.Hiragana => "Hiragana",
            XLPhoneticType.NoConversion => "noConversion",
            _ => throw UnknownValue(value),
        };

    internal static XLPhoneticAlignment ParsePhoneticAlignment(string value) =>
        value switch
        {
            "noControl" => XLPhoneticAlignment.NoControl,
            "left" => XLPhoneticAlignment.Left,
            "center" => XLPhoneticAlignment.Center,
            "distributed" => XLPhoneticAlignment.Distributed,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLPhoneticAlignment value) =>
        value switch
        {
            XLPhoneticAlignment.NoControl => "noControl",
            XLPhoneticAlignment.Left => "left",
            XLPhoneticAlignment.Center => "center",
            XLPhoneticAlignment.Distributed => "distributed",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Sparklines (x14)

    internal static XLSparklineType ParseSparklineType(string value) =>
        value switch
        {
            "line" => XLSparklineType.Line,
            "column" => XLSparklineType.Column,
            "stacked" => XLSparklineType.Stacked,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLSparklineType value) =>
        value switch
        {
            XLSparklineType.Line => "line",
            XLSparklineType.Column => "column",
            XLSparklineType.Stacked => "stacked",
            _ => throw UnknownValue(value),
        };

    internal static XLSparklineAxisMinMax ParseSparklineAxisMinMax(string value) =>
        value switch
        {
            "individual" => XLSparklineAxisMinMax.Automatic,
            "group" => XLSparklineAxisMinMax.SameForAll,
            "custom" => XLSparklineAxisMinMax.Custom,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLSparklineAxisMinMax value) =>
        value switch
        {
            XLSparklineAxisMinMax.Automatic => "individual",
            XLSparklineAxisMinMax.SameForAll => "group",
            XLSparklineAxisMinMax.Custom => "custom",
            _ => throw UnknownValue(value),
        };

    internal static XLDisplayBlanksAsValues ParseDisplayBlanksAs(string value) =>
        value switch
        {
            "span" => XLDisplayBlanksAsValues.Interpolate,
            "gap" => XLDisplayBlanksAsValues.NotPlotted,
            "zero" => XLDisplayBlanksAsValues.Zero,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static string ToXml(this XLDisplayBlanksAsValues value) =>
        value switch
        {
            XLDisplayBlanksAsValues.Interpolate => "span",
            XLDisplayBlanksAsValues.NotPlotted => "gap",
            XLDisplayBlanksAsValues.Zero => "zero",
            _ => throw UnknownValue(value),
        };

    #endregion

    #region Drawings

    /// <summary>
    /// Only written, never read: pictures are anchored by the element the drawing part uses, and
    /// the load path infers the placement from that element rather than from this attribute.
    /// </summary>
    internal static string ToXml(this XLPicturePlacement value) =>
        value switch
        {
            XLPicturePlacement.MoveAndSize => "twoCell",
            XLPicturePlacement.Move => "oneCell",
            XLPicturePlacement.FreeFloating => "absolute",
            _ => throw UnknownValue(value),
        };

    #endregion

    private static ArgumentOutOfRangeException UnknownValue<T>(T value)
        where T : struct, Enum => new(nameof(value), value, $"Unknown {typeof(T).Name} value.");
}

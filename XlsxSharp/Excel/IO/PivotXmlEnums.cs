using XlsxSharp.Excel.PivotValues;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// The enumerations of the pivot table parts, keyed by the string OOXML writes for them.
/// </summary>
/// <remarks>
/// These replace the SDK enum values the pivot readers used to convert from. The strings are the
/// ones the SDK serialises, and the tests check every entry against the conversion it replaces
/// rather than against the schema, so a wrong spelling shows up as a failing test instead of a
/// pivot table that silently loads with the wrong aggregation.
/// </remarks>
internal static class PivotXmlEnums
{
    internal static XLPivotSummary ParseSubtotal(string value) =>
        value switch
        {
            "average" => XLPivotSummary.Average,
            "count" => XLPivotSummary.Count,
            "countNums" => XLPivotSummary.CountNumbers,
            "max" => XLPivotSummary.Maximum,
            "min" => XLPivotSummary.Minimum,
            "product" => XLPivotSummary.Product,
            "stdDev" => XLPivotSummary.StandardDeviation,
            "stdDevp" => XLPivotSummary.PopulationStandardDeviation,
            "sum" => XLPivotSummary.Sum,
            "var" => XLPivotSummary.Variance,
            "varp" => XLPivotSummary.PopulationVariance,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotCalculation ParseShowDataAs(string value) =>
        value switch
        {
            "normal" => XLPivotCalculation.Normal,
            "difference" => XLPivotCalculation.DifferenceFrom,
            "percent" => XLPivotCalculation.PercentageOf,
            "percentDiff" => XLPivotCalculation.PercentageDifferenceFrom,
            "runTotal" => XLPivotCalculation.RunningTotal,
            "percentOfRow" => XLPivotCalculation.PercentageOfRow,
            "percentOfCol" => XLPivotCalculation.PercentageOfColumn,
            "percentOfTotal" => XLPivotCalculation.PercentageOfTotal,
            "index" => XLPivotCalculation.Index,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotAreaType ParsePivotAreaType(string value) =>
        value switch
        {
            "none" => XLPivotAreaType.None,
            "normal" => XLPivotAreaType.Normal,
            "data" => XLPivotAreaType.Data,
            "all" => XLPivotAreaType.All,
            "origin" => XLPivotAreaType.Origin,
            "button" => XLPivotAreaType.Button,
            "topRight" => XLPivotAreaType.TopRight,
            "topEnd" => XLPivotAreaType.TopEnd,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotItemType ParseItemType(string value) =>
        value switch
        {
            "data" => XLPivotItemType.Data,
            "default" => XLPivotItemType.Default,
            "sum" => XLPivotItemType.Sum,
            "countA" => XLPivotItemType.CountA,
            "avg" => XLPivotItemType.Avg,
            "max" => XLPivotItemType.Max,
            "min" => XLPivotItemType.Min,
            "product" => XLPivotItemType.Product,
            "count" => XLPivotItemType.Count,
            "stdDev" => XLPivotItemType.StdDev,
            "stdDevP" => XLPivotItemType.StdDevP,
            "var" => XLPivotItemType.Var,
            "varP" => XLPivotItemType.VarP,
            "grand" => XLPivotItemType.Grand,
            "blank" => XLPivotItemType.Blank,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    /// <summary>
    /// Note that <c>manual</c> is the unsorted state, which the workbook model calls Default.
    /// </summary>
    internal static XLPivotSortType ParseFieldSort(string value) =>
        value switch
        {
            "manual" => XLPivotSortType.Default,
            "ascending" => XLPivotSortType.Ascending,
            "descending" => XLPivotSortType.Descending,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotFormatAction ParseFormatAction(string value) =>
        value switch
        {
            "blank" => XLPivotFormatAction.Blank,
            "formatting" => XLPivotFormatAction.Formatting,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotCfScope ParseCfScope(string value) =>
        value switch
        {
            "selection" => XLPivotCfScope.SelectedCells,
            "data" => XLPivotCfScope.DataFields,
            "field" => XLPivotCfScope.FieldIntersections,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };

    internal static XLPivotCfRuleType ParseCfRuleType(string value) =>
        value switch
        {
            "none" => XLPivotCfRuleType.None,
            "all" => XLPivotCfRuleType.All,
            "row" => XLPivotCfRuleType.Row,
            "column" => XLPivotCfRuleType.Column,
            _ => throw PartStructureException.InvalidAttributeValue(value),
        };
}

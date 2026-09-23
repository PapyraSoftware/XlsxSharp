using System.Xml;
using System.Xml.Schema;
using XlsxSharp.Excel;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.IO;
using XlsxSharp.Excel.IO.Schemas;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.PivotValues;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Tables;
using XlsxSharp.IO;

namespace XlsxSharp.Tests.Excel.IO;

/// <summary>
/// Checks every enumeration XlsxSharp reads or writes against the schema that defines it - the
/// same <c>.xsd</c> files <see cref="SchemaValidator"/> validates saved packages with.
/// </summary>
/// <remarks>
/// <para>
/// Each workbook model value has to write a string the schema allows and read back as itself,
/// and each string the schema allows has to read as the model value of the same name. Where
/// the names differ on purpose, the pairing is spelled out in the test, so a swapped pair or a
/// misspelling fails here rather than loading a worksheet with the wrong page order.
/// </para>
/// <para>
/// Schema values the model has no place for are listed per test as unsupported. Reading one of
/// them is an error today; the list is the inventory of those gaps.
/// </para>
/// </remarks>
public class SchemaEnumerationTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string SharedTypes =
        "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes";

    private const string Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    private const string SpreadsheetDrawing =
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private static readonly XmlToEnumMapper Mapper = XmlToEnumMapper.Instance;

    #region Styles

    [Test]
    public void FontScheme() =>
        AssertMatchesSchema(ValuesOf("ST_FontScheme"), Mapper.Parse<XLFontScheme>, Mapper.GetText);

    [Test]
    public void Underline() =>
        AssertMatchesSchema(
            ValuesOf("ST_UnderlineValues"),
            Mapper.Parse<XLFontUnderlineValues>,
            Mapper.GetText
        );

    [Test]
    public void VerticalTextAlignment() =>
        AssertMatchesSchema(
            ValuesOf("ST_VerticalAlignRun", SharedTypes),
            Mapper.Parse<XLFontVerticalTextAlignmentValues>,
            Mapper.GetText
        );

    [Test]
    public void FillPattern() =>
        AssertMatchesSchema(
            ValuesOf("ST_PatternType"),
            Mapper.Parse<XLFillPatternValues>,
            Mapper.GetText
        );

    [Test]
    public void BorderStyle() =>
        AssertMatchesSchema(
            ValuesOf("ST_BorderStyle"),
            Mapper.Parse<XLBorderStyleValues>,
            Mapper.GetText
        );

    [Test]
    public void HorizontalAlignment() =>
        AssertMatchesSchema(
            ValuesOf("ST_HorizontalAlignment"),
            Mapper.Parse<XLAlignmentHorizontalValues>,
            Mapper.GetText
        );

    [Test]
    public void VerticalAlignment() =>
        AssertMatchesSchema(
            ValuesOf("ST_VerticalAlignment"),
            Mapper.Parse<XLAlignmentVerticalValues>,
            Mapper.GetText
        );

    [Test]
    public void GradientType() =>
        AssertMatchesSchema(
            ValuesOf("ST_GradientType"),
            Mapper.Parse<XLGradientType>,
            Mapper.GetText
        );

    #endregion

    #region Worksheet

    [Test]
    public void PageOrientation() =>
        AssertMatchesSchema(
            ValuesOf("ST_Orientation"),
            WorksheetXmlEnums.ParsePageOrientation,
            v => v.ToXml()
        );

    [Test]
    public void PageOrder() =>
        AssertMatchesSchema(
            ValuesOf("ST_PageOrder"),
            WorksheetXmlEnums.ParsePageOrder,
            v => v.ToXml()
        );

    [Test]
    public void ShowComments() =>
        AssertMatchesSchema(
            ValuesOf("ST_CellComments"),
            WorksheetXmlEnums.ParseShowComments,
            v => v.ToXml()
        );

    [Test]
    public void PrintError() =>
        AssertMatchesSchema(
            ValuesOf("ST_PrintError"),
            WorksheetXmlEnums.ParsePrintError,
            v => v.ToXml()
        );

    [Test]
    public void SheetViewType() =>
        AssertMatchesSchema(
            ValuesOf("ST_SheetViewType"),
            WorksheetXmlEnums.ParseSheetViewType,
            v => v.ToXml()
        );

    [Test]
    public void DataValidationType() =>
        AssertMatchesSchema(
            ValuesOf("ST_DataValidationType"),
            WorksheetXmlEnums.ParseAllowedValues,
            v => v.ToXml(),
            renamed: new Dictionary<XLAllowedValues, string>
            {
                [XLAllowedValues.AnyValue] = "none",
                [XLAllowedValues.WholeNumber] = "whole",
            }
        );

    [Test]
    public void DataValidationErrorStyle() =>
        AssertMatchesSchema(
            ValuesOf("ST_DataValidationErrorStyle"),
            WorksheetXmlEnums.ParseErrorStyle,
            v => v.ToXml()
        );

    [Test]
    public void DataValidationOperator() =>
        AssertMatchesSchema(
            ValuesOf("ST_DataValidationOperator"),
            WorksheetXmlEnums.ParseDataValidationOperator,
            v => v.ToXml(),
            renamed: new Dictionary<XLOperator, string>
            {
                [XLOperator.EqualTo] = "equal",
                [XLOperator.NotEqualTo] = "notEqual",
                [XLOperator.EqualOrLessThan] = "lessThanOrEqual",
                [XLOperator.EqualOrGreaterThan] = "greaterThanOrEqual",
            }
        );

    [Test]
    public void ConditionalFormatType() =>
        AssertMatchesSchema(
            ValuesOf("ST_CfType"),
            WorksheetXmlEnums.ParseConditionalFormatType,
            v => v.ToXml(),
            renamed: new Dictionary<XLConditionalFormatType, string>
            {
                [XLConditionalFormatType.IsUnique] = "uniqueValues",
                [XLConditionalFormatType.IsDuplicate] = "duplicateValues",
                [XLConditionalFormatType.StartsWith] = "beginsWith",
                [XLConditionalFormatType.IsBlank] = "containsBlanks",
                [XLConditionalFormatType.NotBlank] = "notContainsBlanks",
                [XLConditionalFormatType.IsError] = "containsErrors",
                [XLConditionalFormatType.NotError] = "notContainsErrors",
            }
        );

    [Test]
    public void ConditionalFormatOperator() =>
        AssertMatchesSchema(
            ValuesOf("ST_ConditionalFormattingOperator"),
            WorksheetXmlEnums.ParseCfOperator,
            v => v.ToXml(),
            renamed: new Dictionary<XLCFOperator, string>
            {
                [XLCFOperator.EqualOrLessThan] = "lessThanOrEqual",
                [XLCFOperator.EqualOrGreaterThan] = "greaterThanOrEqual",
                [XLCFOperator.Contains] = "containsText",
                [XLCFOperator.StartsWith] = "beginsWith",
            }
        );

    [Test]
    public void ConditionalFormatValueObjectType() =>
        AssertMatchesSchema(
            ValuesOf("ST_CfvoType"),
            WorksheetXmlEnums.ParseCfContentType,
            v => v.ToXml(),
            renamed: new Dictionary<XLCFContentType, string>
            {
                [XLCFContentType.Number] = "num",
                [XLCFContentType.Maximum] = "max",
                [XLCFContentType.Minimum] = "min",
            }
        );

    [Test]
    public void IconSet() =>
        AssertMatchesSchema(
            ValuesOf("ST_IconSetType"),
            WorksheetXmlEnums.ParseIconSetStyle,
            v => v.ToXml()
        );

    [Test]
    public void TimePeriod() =>
        AssertMatchesSchema(
            ValuesOf("ST_TimePeriod"),
            WorksheetXmlEnums.ParseTimePeriod,
            v => v.ToXml(),
            renamed: new Dictionary<XLTimePeriod, string>
            {
                [XLTimePeriod.InTheLast7Days] = "last7Days",
            }
        );

    [Test]
    public void FilterOperator() =>
        AssertMatchesSchema(
            ValuesOf("ST_FilterOperator"),
            WorksheetXmlEnums.ParseFilterOperator,
            v => v.ToXml(),
            renamed: new Dictionary<XLFilterOperator, string>
            {
                [XLFilterOperator.EqualOrLessThan] = "lessThanOrEqual",
                [XLFilterOperator.EqualOrGreaterThan] = "greaterThanOrEqual",
            }
        );

    [Test]
    public void DynamicFilterType() =>
        AssertMatchesSchema(
            ValuesOf("ST_DynamicFilterType"),
            WorksheetXmlEnums.ParseFilterDynamicType,
            v => v.ToXml(),
            // The model only knows the two average filters, not the date periods.
            unsupported:
            [
                "null",
                "tomorrow",
                "today",
                "yesterday",
                "nextWeek",
                "thisWeek",
                "lastWeek",
                "nextMonth",
                "thisMonth",
                "lastMonth",
                "nextQuarter",
                "thisQuarter",
                "lastQuarter",
                "nextYear",
                "thisYear",
                "lastYear",
                "yearToDate",
                "Q1",
                "Q2",
                "Q3",
                "Q4",
                "M1",
                "M2",
                "M3",
                "M4",
                "M5",
                "M6",
                "M7",
                "M8",
                "M9",
                "M10",
                "M11",
                "M12",
            ]
        );

    [Test]
    public void DateTimeGrouping() =>
        AssertMatchesSchema(
            ValuesOf("ST_DateTimeGrouping"),
            WorksheetXmlEnums.ParseDateTimeGrouping,
            v => v.ToXml()
        );

    [Test]
    public void TotalsRowFunction() =>
        AssertMatchesSchema(
            ValuesOf("ST_TotalsRowFunction"),
            WorksheetXmlEnums.ParseTotalsRowFunction,
            v => v.ToXml(),
            renamed: new Dictionary<XLTotalsRowFunction, string>
            {
                [XLTotalsRowFunction.Minimum] = "min",
                [XLTotalsRowFunction.Maximum] = "max",
                [XLTotalsRowFunction.CountNumbers] = "countNums",
                [XLTotalsRowFunction.StandardDeviation] = "stdDev",
                [XLTotalsRowFunction.Variance] = "var",
            }
        );

    [Test]
    public void PhoneticType() =>
        AssertMatchesSchema(
            ValuesOf("ST_PhoneticType"),
            WorksheetXmlEnums.ParsePhoneticType,
            v => v.ToXml()
        );

    [Test]
    public void PhoneticAlignment() =>
        AssertMatchesSchema(
            ValuesOf("ST_PhoneticAlignment"),
            WorksheetXmlEnums.ParsePhoneticAlignment,
            v => v.ToXml()
        );

    /// <summary>
    /// Sparklines live in the Excel 2010 extension namespace, whose schema is published in
    /// [MS-XLSX] 2.6.x rather than in ECMA-376 and is not among the vendored ones.
    /// </summary>
    [Test]
    public void SparklineType() =>
        AssertMatchesSchema(
            ["line", "column", "stacked"],
            WorksheetXmlEnums.ParseSparklineType,
            v => v.ToXml()
        );

    /// <inheritdoc cref="SparklineType"/>
    [Test]
    public void SparklineAxisMinMax() =>
        AssertMatchesSchema(
            ["individual", "group", "custom"],
            WorksheetXmlEnums.ParseSparklineAxisMinMax,
            v => v.ToXml(),
            renamed: new Dictionary<XLSparklineAxisMinMax, string>
            {
                [XLSparklineAxisMinMax.Automatic] = "individual",
                [XLSparklineAxisMinMax.SameForAll] = "group",
            }
        );

    /// <summary>The x14 sparkline group borrows the chart schema's type for this one.</summary>
    [Test]
    public void DisplayBlanksAs() =>
        AssertMatchesSchema(
            ValuesOf("ST_DispBlanksAs", Chart),
            WorksheetXmlEnums.ParseDisplayBlanksAs,
            v => v.ToXml(),
            renamed: new Dictionary<XLDisplayBlanksAsValues, string>
            {
                [XLDisplayBlanksAsValues.Interpolate] = "span",
                [XLDisplayBlanksAsValues.NotPlotted] = "gap",
            }
        );

    /// <summary>
    /// The placement is only ever written - the load path infers it from the anchor element
    /// itself - so there is nothing to parse.
    /// </summary>
    [Test]
    public void PicturePlacement() =>
        AssertMatchesSchema(
            ValuesOf("ST_EditAs", SpreadsheetDrawing),
            parse: null,
            v => v.ToXml(),
            renamed: new Dictionary<XLPicturePlacement, string>
            {
                [XLPicturePlacement.MoveAndSize] = "twoCell",
                [XLPicturePlacement.Move] = "oneCell",
                [XLPicturePlacement.FreeFloating] = "absolute",
            }
        );

    [Test]
    public void AnUnknownWorksheetValueIsRejected() =>
        ClassicAssert.Throws<PartStructureException>(() =>
            WorksheetXmlEnums.ParsePageOrder("sidewaysThenUp")
        );

    #endregion

    #region Pivot tables

    [Test]
    public void PivotSubtotal() =>
        AssertMatchesSchema(
            ValuesOf("ST_DataConsolidateFunction"),
            PivotXmlEnums.ParseSubtotal,
            toXml: null,
            renamed: new Dictionary<XLPivotSummary, string>
            {
                [XLPivotSummary.CountNumbers] = "countNums",
                [XLPivotSummary.Maximum] = "max",
                [XLPivotSummary.Minimum] = "min",
                [XLPivotSummary.StandardDeviation] = "stdDev",
                [XLPivotSummary.PopulationStandardDeviation] = "stdDevp",
                [XLPivotSummary.Variance] = "var",
                [XLPivotSummary.PopulationVariance] = "varp",
            }
        );

    [Test]
    public void PivotShowDataAs() =>
        AssertMatchesSchema(
            ValuesOf("ST_ShowDataAs"),
            PivotXmlEnums.ParseShowDataAs,
            toXml: null,
            renamed: new Dictionary<XLPivotCalculation, string>
            {
                [XLPivotCalculation.DifferenceFrom] = "difference",
                [XLPivotCalculation.PercentageOf] = "percent",
                [XLPivotCalculation.PercentageDifferenceFrom] = "percentDiff",
                [XLPivotCalculation.RunningTotal] = "runTotal",
                [XLPivotCalculation.PercentageOfRow] = "percentOfRow",
                [XLPivotCalculation.PercentageOfColumn] = "percentOfCol",
                [XLPivotCalculation.PercentageOfTotal] = "percentOfTotal",
            }
        );

    [Test]
    public void PivotAreaType() =>
        AssertMatchesSchema(
            ValuesOf("ST_PivotAreaType"),
            PivotXmlEnums.ParsePivotAreaType,
            toXml: null
        );

    [Test]
    public void PivotItemType() =>
        AssertMatchesSchema(ValuesOf("ST_ItemType"), PivotXmlEnums.ParseItemType, toXml: null);

    [Test]
    public void PivotFieldSort() =>
        AssertMatchesSchema(
            ValuesOf("ST_FieldSortType"),
            PivotXmlEnums.ParseFieldSort,
            toXml: null,
            renamed: new Dictionary<XLPivotSortType, string>
            {
                [XLPivotSortType.Default] = "manual",
            }
        );

    [Test]
    public void PivotFormatAction() =>
        AssertMatchesSchema(
            ValuesOf("ST_FormatAction"),
            PivotXmlEnums.ParseFormatAction,
            toXml: null,
            // Both are reserved by the specification and never written by Excel.
            unsupported: ["drill", "formula"]
        );

    [Test]
    public void PivotConditionalFormatScope() =>
        AssertMatchesSchema(
            ValuesOf("ST_Scope"),
            PivotXmlEnums.ParseCfScope,
            toXml: null,
            renamed: new Dictionary<XLPivotCfScope, string>
            {
                [XLPivotCfScope.SelectedCells] = "selection",
                [XLPivotCfScope.DataFields] = "data",
                [XLPivotCfScope.FieldIntersections] = "field",
            }
        );

    [Test]
    public void PivotConditionalFormatRuleType() =>
        AssertMatchesSchema(ValuesOf("ST_Type"), PivotXmlEnums.ParseCfRuleType, toXml: null);

    [Test]
    public void PivotAxis() =>
        AssertMatchesSchema(ValuesOf("ST_Axis"), PivotXmlEnums.ParseAxis, toXml: null);

    [Test]
    public void AnUnknownPivotValueIsRejected() =>
        ClassicAssert.Throws<PartStructureException>(() =>
            PivotXmlEnums.ParseSubtotal("notAFunction")
        );

    #endregion

    /// <summary>
    /// The enumeration values of a simple type in the vendored schemas, in schema order.
    /// </summary>
    private static string[] ValuesOf(string simpleTypeName, string ns = Main)
    {
        XmlSchemaSimpleType type =
            OoxmlSchemas.Set.GlobalTypes[new XmlQualifiedName(simpleTypeName, ns)]
                as XmlSchemaSimpleType
            ?? throw new ArgumentException($"No simple type {simpleTypeName} in {ns}.");

        string[] values =
        [
            .. ((XmlSchemaSimpleTypeRestriction)type.Content!)
                .Facets.OfType<XmlSchemaEnumerationFacet>()
                .Select(facet => facet.Value!),
        ];

        ClassicAssert.IsTrue(values.Length > 0, simpleTypeName);
        return values;
    }

    /// <param name="schemaValues">Every value the schema allows.</param>
    /// <param name="parse">Reads a string, or <c>null</c> for an enumeration only ever written.</param>
    /// <param name="toXml">Writes a value, or <c>null</c> for an enumeration only ever read.</param>
    /// <param name="renamed">
    /// The model values whose name is not the schema's string, with the string they stand for.
    /// </param>
    /// <param name="unsupported">Schema values the model has no value for.</param>
    private static void AssertMatchesSchema<TEnum>(
        IReadOnlyCollection<string> schemaValues,
        Func<string, TEnum>? parse,
        Func<TEnum, string>? toXml,
        IReadOnlyDictionary<TEnum, string>? renamed = null,
        IReadOnlyCollection<string>? unsupported = null
    )
        where TEnum : struct, Enum
    {
        unsupported ??= [];
        ClassicAssert.IsTrue(
            unsupported.All(schemaValues.Contains),
            "unsupported must be schema values"
        );

        string Expected(TEnum value) =>
            renamed is not null && renamed.TryGetValue(value, out string? text)
                ? text
                : value.ToString();

        List<string> mismatches = [];
        if (parse is not null)
        {
            foreach (string schemaValue in schemaValues)
            {
                if (unsupported.Contains(schemaValue))
                {
                    ClassicAssert.Throws<PartStructureException>(
                        () => parse(schemaValue),
                        $"'{schemaValue}' is listed as unsupported, but reads"
                    );
                    continue;
                }

                TEnum parsed = parse(schemaValue);
                if (!SameName(Expected(parsed), schemaValue))
                {
                    mismatches.Add($"[{typeof(TEnum).Name}.{parsed}] = \"{schemaValue}\",");
                }
            }
        }

        if (toXml is not null)
        {
            foreach (TEnum value in Enum.GetValues<TEnum>())
            {
                string written = toXml(value);
                ClassicAssert.IsTrue(
                    schemaValues.Contains(written),
                    $"{typeof(TEnum).Name}.{value} writes '{written}', which the schema does not allow"
                );
                if (!SameName(Expected(value), written))
                {
                    mismatches.Add($"[{typeof(TEnum).Name}.{value}] = \"{written}\",");
                }

                if (parse is not null)
                {
                    ClassicAssert.AreEqual(
                        value,
                        parse(written),
                        $"'{written}' does not read back"
                    );
                }
            }
        }

        // Every pairing whose names differ, so that a failure lists all of them at once.
        ClassicAssert.IsTrue(
            mismatches.Count == 0,
            "Names differ, pair them in renamed:"
                + Environment.NewLine
                + string.Join(Environment.NewLine, mismatches.Distinct())
        );
    }

    /// <summary>
    /// The model's names are the schema's in PascalCase, give or take a separator:
    /// <c>DownThenOver</c> for <c>downThenOver</c>, <c>ThreeArrows</c> for <c>3Arrows</c>.
    /// </summary>
    private static bool SameName(string modelName, string schemaValue) =>
        Normalize(modelName) == Normalize(schemaValue);

    private static string Normalize(string name) =>
        new string([.. name.Where(char.IsLetterOrDigit)])
            .ToLowerInvariant()
            .Replace("three", "3")
            .Replace("four", "4")
            .Replace("five", "5");
}

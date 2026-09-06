using System.Globalization;
using System.Xml.Linq;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writing the <c>cfRule</c> of one conditional format.
/// </summary>
/// <remarks>
/// Every rule is the same element with the same three optional attributes; the seventeen kinds
/// differ in which of them they carry and in the formula they write beside it. Excel evaluates
/// the formula rather than the kind, so a rule that says "is blank" also has to say
/// <c>LEN(TRIM(A2))=0</c>, and the formulas are what most of the kinds amount to.
/// </remarks>
internal static class ConditionalFormatXml
{
    /// <summary>
    /// The formulas Excel writes beside a "dates occurring" rule, one per period.
    /// </summary>
    private static readonly Dictionary<XLTimePeriod, string> TimePeriodFormulas = new()
    {
        [XLTimePeriod.Today] = "FLOOR({0},1)=TODAY()",
        [XLTimePeriod.Yesterday] = "FLOOR({0},1)=TODAY()-1",
        [XLTimePeriod.Tomorrow] = "FLOOR({0},1)=TODAY()+1",
        [XLTimePeriod.InTheLast7Days] = "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
        [XLTimePeriod.ThisMonth] = "AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))",
        [XLTimePeriod.LastMonth] =
            "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))",
        [XLTimePeriod.NextMonth] =
            "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)),YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
        [XLTimePeriod.ThisWeek] =
            "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
        [XLTimePeriod.LastWeek] =
            "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
        [XLTimePeriod.NextWeek] =
            "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))",
    };

    internal static XElement Rule(XLConditionalFormat cf, int priority, SaveContext context)
    {
        XElement rule = new(
            SpreadsheetXml.Main + "cfRule",
            new XAttribute("type", cf.ConditionalFormatType.ToXml()),
            new XAttribute("priority", priority)
        );
        WorksheetXml.SetBoolDefault(rule, "stopIfTrue", cf.StopIfTrue, false);

        // The three kinds that paint the cell themselves carry no differential format.
        bool paintsItself =
            cf.ConditionalFormatType
            is XLConditionalFormatType.ColorScale
                or XLConditionalFormatType.DataBar
                or XLConditionalFormatType.IconSet;
        if (!paintsItself && cf.FormatValue is { } format)
        {
            rule.SetAttributeValue("dxfId", context.GetDxfId(format));
        }

        string firstCell = cf.Range.RangeAddress.FirstAddress.ToStringRelative(false);
        switch (cf.ConditionalFormatType)
        {
            case XLConditionalFormatType.CellIs:
            case XLConditionalFormatType.Expression:
                WriteCellIs(rule, cf);
                break;

            case XLConditionalFormatType.ContainsText:
                WriteTextMatch(
                    rule,
                    cf,
                    "containsText",
                    value => $"NOT(ISERROR(SEARCH(\"{value}\",{firstCell})))"
                );
                break;

            case XLConditionalFormatType.NotContainsText:
                WriteTextMatch(
                    rule,
                    cf,
                    "notContains",
                    value => $"ISERROR(SEARCH(\"{value}\",{firstCell}))"
                );
                break;

            case XLConditionalFormatType.StartsWith:
                WriteTextMatch(
                    rule,
                    cf,
                    "beginsWith",
                    value => $"LEFT({firstCell},{value.Length})=\"{value}\""
                );
                break;

            case XLConditionalFormatType.EndsWith:
                WriteTextMatch(
                    rule,
                    cf,
                    "endsWith",
                    value => $"RIGHT({firstCell},{value.Length})=\"{value}\""
                );
                break;

            case XLConditionalFormatType.IsBlank:
                rule.Add(Formula($"LEN(TRIM({firstCell}))=0"));
                break;

            case XLConditionalFormatType.NotBlank:
                rule.Add(Formula($"LEN(TRIM({firstCell}))>0"));
                break;

            case XLConditionalFormatType.IsError:
                rule.Add(Formula($"ISERROR({firstCell})"));
                break;

            case XLConditionalFormatType.NotError:
                rule.Add(Formula($"NOT(ISERROR({firstCell}))"));
                break;

            case XLConditionalFormatType.TimePeriod:
                rule.SetAttributeValue("timePeriod", cf.TimePeriod.ToXml());
                rule.Add(
                    Formula(
                        string.Format(
                            CultureInfo.InvariantCulture,
                            TimePeriodFormulas[cf.TimePeriod],
                            firstCell
                        )
                    )
                );
                break;

            case XLConditionalFormatType.Top10:
                WorksheetXml.SetBool(rule, "percent", cf.Percent);
                WorksheetXml.Set(
                    rule,
                    "rank",
                    uint.Parse(cf.Values[1].Value, CultureInfo.InvariantCulture)
                );
                WorksheetXml.SetBool(rule, "bottom", cf.Bottom);
                break;

            case XLConditionalFormatType.ColorScale:
                rule.Add(WriteColorScale(cf));
                break;

            case XLConditionalFormatType.DataBar:
                rule.Add(WriteDataBar(cf), WriteDataBarExtension(cf));
                break;

            case XLConditionalFormatType.IconSet:
                rule.Add(WriteIconSet(cf));
                break;

            case XLConditionalFormatType.IsUnique:
            case XLConditionalFormatType.IsDuplicate:
                // The kind is the whole rule.
                break;

            default:
                throw new NotImplementedException(
                    $"Conditional formatting rule '{cf.ConditionalFormatType}' hasn't been implemented"
                );
        }

        return rule;
    }

    private static void WriteCellIs(XElement rule, XLConditionalFormat cf)
    {
        rule.SetAttributeValue("operator", cf.Operator.ToXml());
        rule.Add(Formula(Quoted(cf.Values[1])));

        if (cf.Operator is XLCFOperator.Between or XLCFOperator.NotBetween)
        {
            rule.Add(Formula(Quoted(cf.Values[2])));
        }
    }

    private static void WriteTextMatch(
        XElement rule,
        XLConditionalFormat cf,
        string operatorName,
        Func<string, string> formula
    )
    {
        string value = cf.Values[1].Value;
        rule.SetAttributeValue("operator", operatorName);
        rule.SetAttributeValue("text", value);
        rule.Add(Formula(formula(value)));
    }

    private static XElement WriteColorScale(XLConditionalFormat cf)
    {
        XElement colorScale = new(SpreadsheetXml.Main + "colorScale");
        for (int i = 1; i <= cf.ContentTypes.Count; i++)
        {
            colorScale.Add(
                ValueObject(
                    cf.ContentTypes[i],
                    cf.Values.TryGetValue(i, out XLFormula? formula) ? formula?.Value : null
                )
            );
        }

        for (int i = 1; i <= cf.Colors.Count; i++)
        {
            colorScale.Add(Color(cf.Colors[i]));
        }

        return colorScale;
    }

    private static XElement WriteDataBar(XLConditionalFormat cf)
    {
        XElement dataBar = new(SpreadsheetXml.Main + "dataBar");
        WorksheetXml.SetBool(dataBar, "showValue", !cf.ShowBarOnly);
        dataBar.Add(
            BoundValueObject(cf, 1, XLCFContentType.Minimum),
            BoundValueObject(cf, 2, XLCFContentType.Maximum),
            Color(cf.Colors[1])
        );
        return dataBar;
    }

    /// <summary>
    /// The guid that ties this rule to the x14 rule carrying the parts of a data bar the 2006
    /// schema cannot express.
    /// </summary>
    private static XElement WriteDataBarExtension(XLConditionalFormat cf) =>
        new(
            SpreadsheetXml.Main + "extLst",
            new XElement(
                SpreadsheetXml.Main + "ext",
                new XAttribute(XNamespace.Xmlns + "x14", SpreadsheetXml.X14.NamespaceName),
                new XAttribute("uri", "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"),
                new XElement(SpreadsheetXml.X14 + "id", cf.Id.WrapInBraces())
            )
        );

    private static XElement WriteIconSet(XLConditionalFormat cf)
    {
        XElement iconSet = new(
            SpreadsheetXml.Main + "iconSet",
            new XAttribute("iconSet", cf.IconSetStyle.ToXml())
        );
        WorksheetXml.SetBool(iconSet, "showValue", !cf.ShowIconOnly);
        WorksheetXml.SetBool(iconSet, "reverse", cf.ReverseIconOrder);

        for (int i = 1; i <= cf.Values.Count; i++)
        {
            XElement valueObject = ValueObject(cf.ContentTypes[i], cf.Values[i].Value);
            WorksheetXml.SetBool(
                valueObject,
                "gte",
                cf.IconSetOperators[i] == XLCFIconSetOperator.EqualOrGreaterThan
            );
            iconSet.Add(valueObject);
        }

        return iconSet;
    }

    /// <summary>
    /// The x14 half of a data bar, which is where its negative colour, its axis and its lengths
    /// live. It names the guid of the rule it belongs to.
    /// </summary>
    internal static XElement ExtensionRule(XLConditionalFormat cf)
    {
        XElement dataBar = new(
            SpreadsheetXml.X14 + "dataBar",
            new XAttribute("minLength", 0),
            new XAttribute("maxLength", 100)
        );
        WorksheetXml.SetBool(dataBar, "showValue", !cf.ShowBarOnly);
        WorksheetXml.SetBool(dataBar, "gradient", true);

        dataBar.Add(ExtensionValueObject(cf, 1, "autoMin"), ExtensionValueObject(cf, 2, "autoMax"));

        // A second colour is the one negative bars are painted with; with only one, the bars are
        // painted the same colour in both directions.
        XLColor negative = cf.Colors.Count == 2 ? cf.Colors[2] : cf.Colors[1];
        XElement negativeFillColor = new(SpreadsheetXml.X14 + "negativeFillColor");
        SpreadsheetXml.SetColor(negativeFillColor, negative);
        XElement axisColor = new(SpreadsheetXml.X14 + "axisColor");
        SpreadsheetXml.SetColor(axisColor, XLColor.Black);
        dataBar.Add(negativeFillColor, axisColor);

        return new XElement(
            SpreadsheetXml.X14 + "cfRule",
            new XAttribute("type", XLConditionalFormatType.DataBar.ToXml()),
            new XAttribute("id", cf.Id.WrapInBraces()),
            dataBar
        );
    }

    /// <summary>
    /// A bound of the x14 data bar. A bound with a value of its own is numeric whatever the
    /// workbook model calls it, and the value is a formula element rather than an attribute.
    /// </summary>
    private static XElement ExtensionValueObject(
        XLConditionalFormat cf,
        int index,
        string automatic
    )
    {
        string type = cf.ContentTypes.TryGetValue(index, out XLCFContentType contentType)
            ? ExtensionType(contentType)
            : automatic;

        XElement valueObject = new(SpreadsheetXml.X14 + "cfvo");
        if (cf.Values.Count >= index && cf.Values[index]?.Value is { } value)
        {
            type = "num";
            valueObject.Add(new XElement(SpreadsheetXml.Xm + "f", value));
        }

        valueObject.SetAttributeValue("type", type);
        return valueObject;
    }

    /// <summary>
    /// The x14 bound types, which say <c>num</c> where the 2006 schema says <c>num</c> but
    /// <c>autoMin</c> and <c>autoMax</c> where it says <c>min</c> and <c>max</c>.
    /// </summary>
    private static string ExtensionType(XLCFContentType contentType) =>
        contentType switch
        {
            XLCFContentType.Minimum => "autoMin",
            XLCFContentType.Maximum => "autoMax",
            XLCFContentType.Number => "num",
            XLCFContentType.Percent => "percent",
            XLCFContentType.Percentile => "percentile",
            XLCFContentType.Formula => "formula",
            _ => throw new ArgumentOutOfRangeException(nameof(contentType), contentType, null),
        };

    private static XElement BoundValueObject(
        XLConditionalFormat cf,
        int index,
        XLCFContentType defaultType
    ) =>
        ValueObject(
            cf.ContentTypes.TryGetValue(index, out XLCFContentType contentType)
                ? contentType
                : defaultType,
            cf.Values.TryGetValue(index, out XLFormula? value) ? value?.Value : null
        );

    private static XElement ValueObject(XLCFContentType type, string? value)
    {
        XElement valueObject = new(
            SpreadsheetXml.Main + "cfvo",
            new XAttribute("type", type.ToXml())
        );
        valueObject.SetAttributeValue("val", value);
        return valueObject;
    }

    private static XElement Color(XLColor xlColor)
    {
        XElement color = new(SpreadsheetXml.Main + "color");
        switch (xlColor.ColorType)
        {
            case XLColorType.Color:
                color.SetAttributeValue("rgb", xlColor.Color.ToHex());
                break;
            case XLColorType.Theme:
                color.SetAttributeValue("theme", (uint)xlColor.ThemeColor);
                break;
            case XLColorType.Indexed:
                color.SetAttributeValue("indexed", (uint)xlColor.Indexed);
                break;
        }

        return color;
    }

    private static XElement Formula(string text) => new(SpreadsheetXml.Main + "formula", text);

    /// <summary>
    /// A cellIs argument that is neither a formula nor a number is a string, and a string has to
    /// be quoted to be one - with its own quotes doubled.
    /// </summary>
    private static string Quoted(XLFormula formula)
    {
        string value = formula.Value;

        if (
            formula.IsFormula
            || (value.StartsWith('"') && value.EndsWith('"'))
            || double.TryParse(
                value,
                XlsxSharp.XLHelper.NumberStyle,
                XlsxSharp.XLHelper.ParseCulture,
                out _
            )
        )
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}

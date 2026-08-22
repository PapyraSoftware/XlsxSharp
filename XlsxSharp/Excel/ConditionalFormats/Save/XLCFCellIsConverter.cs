using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Misc;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFCellIsConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        string val = GetQuoted(cf.Values[1]);

        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        conditionalFormattingRule.Operator = cf.Operator.ToOpenXml();

        Formula formula = new(val);
        conditionalFormattingRule.Append(formula);

        if (cf.Operator == XLCFOperator.Between || cf.Operator == XLCFOperator.NotBetween)
        {
            Formula formula2 = new() { Text = GetQuoted(cf.Values[2]) };
            conditionalFormattingRule.Append(formula2);
        }

        return conditionalFormattingRule;
    }

    private static string GetQuoted(XLFormula formula)
    {
        string value = formula.Value;

        if (
            formula.IsFormula
            || value.StartsWith("\"") && value.EndsWith("\"")
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

        return string.Format("\"{0}\"", value.Replace("\"", "\"\""));
    }
}

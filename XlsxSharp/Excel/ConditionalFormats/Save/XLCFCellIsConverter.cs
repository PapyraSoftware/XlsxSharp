using System;
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
        String val = GetQuoted(cf.Values[1]);

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

    private static String GetQuoted(XLFormula formula)
    {
        String value = formula.Value;

        if (
            formula.IsFormula
            || value.StartsWith("\"") && value.EndsWith("\"")
            || Double.TryParse(
                value,
                XlsxSharp.XLHelper.NumberStyle,
                XlsxSharp.XLHelper.ParseCulture,
                out _
            )
        )
        {
            return value;
        }

        return String.Format("\"{0}\"", value.Replace("\"", "\"\""));
    }
}

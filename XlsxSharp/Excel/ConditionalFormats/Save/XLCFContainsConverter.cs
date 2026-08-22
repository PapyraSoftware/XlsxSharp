using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFContainsConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        string val = cf.Values[1].Value;
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.ContainsText;
        conditionalFormattingRule.Text = val;

        Formula formula = new()
        {
            Text =
                "NOT(ISERROR(SEARCH(\""
                + val
                + "\","
                + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false)
                + ")))",
        };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}

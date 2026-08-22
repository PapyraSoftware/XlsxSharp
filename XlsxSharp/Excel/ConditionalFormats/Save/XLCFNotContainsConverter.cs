using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFNotContainsConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        String val = cf.Values[1].Value;
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.NotContains;
        conditionalFormattingRule.Text = val;

        Formula formula = new()
        {
            Text =
                "ISERROR(SEARCH(\""
                + val
                + "\","
                + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false)
                + "))",
        };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}

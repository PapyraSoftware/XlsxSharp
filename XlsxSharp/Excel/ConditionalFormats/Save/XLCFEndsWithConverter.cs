using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFEndsWithConverter : IXLCFConverter
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
        conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.EndsWith;
        conditionalFormattingRule.Text = val;

        Formula formula = new()
        {
            Text =
                "RIGHT("
                + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false)
                + ","
                + val.Length.ToString()
                + ")=\""
                + val
                + "\"",
        };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}

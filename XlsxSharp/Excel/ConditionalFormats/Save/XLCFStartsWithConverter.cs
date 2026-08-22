using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFStartsWithConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        string? val = cf.Values[1].Value;
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.BeginsWith;
        conditionalFormattingRule.Text = val;

        Formula formula = new()
        {
            Text =
                "LEFT("
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

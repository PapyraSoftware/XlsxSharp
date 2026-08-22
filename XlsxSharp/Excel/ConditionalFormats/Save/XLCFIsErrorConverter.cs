using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFIsErrorConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        Formula formula = new()
        {
            Text = "ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")",
        };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}

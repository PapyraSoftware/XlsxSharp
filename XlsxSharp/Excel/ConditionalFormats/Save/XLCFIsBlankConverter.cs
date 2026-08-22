using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFIsBlankConverter : IXLCFConverter
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
            Text =
                "LEN(TRIM(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))=0",
        };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}

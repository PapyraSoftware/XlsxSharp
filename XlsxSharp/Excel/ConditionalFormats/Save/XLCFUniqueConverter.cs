using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFUniqueConverter : IXLCFConverter
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
        return conditionalFormattingRule;
    }
}

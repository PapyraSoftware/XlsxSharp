using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Utils;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal static class XLCFBaseConverter
{
    public static ConditionalFormattingRule Convert(XLConditionalFormat cf, int priority) =>
        new()
        {
            Type = cf.ConditionalFormatType.ToOpenXml(),
            Priority = priority,
            StopIfTrue = OpenXmlHelper.GetBooleanValue(cf.StopIfTrue, false),
        };

    public static ConditionalFormattingRule ConvertWithDxf(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        ConditionalFormattingRule cfRule = Convert(cf, priority);
        cfRule.FormatId = cf.FormatValue is not null ? context.GetDxfId(cf.FormatValue) : null;
        return cfRule;
    }
}

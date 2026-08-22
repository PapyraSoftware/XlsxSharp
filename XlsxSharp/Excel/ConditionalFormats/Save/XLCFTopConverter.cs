using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFTopConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        uint val = uint.Parse(cf.Values[1].Value);
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.ConvertWithDxf(
            cf,
            priority,
            context
        );
        conditionalFormattingRule.Percent = cf.Percent;
        conditionalFormattingRule.Rank = val;
        conditionalFormattingRule.Bottom = cf.Bottom;
        return conditionalFormattingRule;
    }
}

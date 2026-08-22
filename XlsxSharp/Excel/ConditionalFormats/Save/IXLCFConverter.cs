using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal interface IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    );
}

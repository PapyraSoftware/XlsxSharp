#nullable disable

using DocumentFormat.OpenXml.Office2010.Excel;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal interface IXLCFConverterExtension
{
    ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context);
}

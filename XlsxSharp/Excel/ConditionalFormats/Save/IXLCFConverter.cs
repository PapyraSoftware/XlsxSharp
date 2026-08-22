using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal interface IXLCFConverter
{
    ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        Int32 priority,
        XLWorkbook.SaveContext context
    );
}

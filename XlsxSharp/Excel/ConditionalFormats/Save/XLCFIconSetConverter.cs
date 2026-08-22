#nullable disable

using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFIconSetConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(
        XLConditionalFormat cf,
        int priority,
        XLWorkbook.SaveContext context
    )
    {
        ConditionalFormattingRule conditionalFormattingRule = XLCFBaseConverter.Convert(
            cf,
            priority
        );

        IconSet iconSet = new()
        {
            ShowValue = !cf.ShowIconOnly,
            Reverse = cf.ReverseIconOrder,
            IconSetValue = cf.IconSetStyle.ToOpenXml(),
        };
        Int32 count = cf.Values.Count;
        for (Int32 i = 1; i <= count; i++)
        {
            ConditionalFormatValueObject conditionalFormatValueObject = new()
            {
                Type = cf.ContentTypes[i].ToOpenXml(),
                Val = cf.Values[i].Value,
                GreaterThanOrEqual =
                    cf.IconSetOperators[i] == XLCFIconSetOperator.EqualOrGreaterThan,
            };
            iconSet.Append(conditionalFormatValueObject);
        }
        conditionalFormattingRule.Append(iconSet);
        return conditionalFormattingRule;
    }
}

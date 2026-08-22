#nullable disable

using System;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFColorScaleConverter : IXLCFConverter
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

        ColorScale colorScale = new();
        for (Int32 i = 1; i <= cf.ContentTypes.Count; i++)
        {
            ConditionalFormatValueObjectValues type = cf.ContentTypes[i].ToOpenXml();
            string val = cf.Values.TryGetValue(i, out XLFormula formula) ? formula?.Value : null;

            ConditionalFormatValueObject conditionalFormatValueObject = new() { Type = type };
            if (val != null)
            {
                conditionalFormatValueObject.Val = val;
            }

            colorScale.Append(conditionalFormatValueObject);
        }

        for (Int32 i = 1; i <= cf.Colors.Count; i++)
        {
            XLColor xlColor = cf.Colors[i];
            Color color = new();
            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    color.Rgb = xlColor.Color.ToHex();
                    break;
                case XLColorType.Theme:
                    color.Theme = System.Convert.ToUInt32(xlColor.ThemeColor);
                    break;

                case XLColorType.Indexed:
                    color.Indexed = System.Convert.ToUInt32(xlColor.Indexed);
                    break;
            }

            colorScale.Append(color);
        }

        conditionalFormattingRule.Append(colorScale);

        return conditionalFormattingRule;
    }
}

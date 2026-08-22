using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Extensions;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using ConditionalFormattingRule = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule;
using DataBar = DocumentFormat.OpenXml.Spreadsheet.DataBar;

namespace XlsxSharp.Excel.ConditionalFormats.Save;

internal class XLCFDataBarConverter : IXLCFConverter
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

        DataBar dataBar = new() { ShowValue = !cf.ShowBarOnly };

        ConditionalFormatValueObject conditionalFormatValueObject1 =
            GetConditionalFormatValueObjectByIndex(cf, 1, ConditionalFormatValueObjectValues.Min);
        ConditionalFormatValueObject conditionalFormatValueObject2 =
            GetConditionalFormatValueObjectByIndex(cf, 2, ConditionalFormatValueObjectValues.Max);

        Color color = new();
        switch (cf.Colors[1].ColorType)
        {
            case XLColorType.Color:
                color.Rgb = cf.Colors[1].Color.ToHex();
                break;

            case XLColorType.Theme:
                color.Theme = System.Convert.ToUInt32(cf.Colors[1].ThemeColor);
                break;

            case XLColorType.Indexed:
                color.Indexed = System.Convert.ToUInt32(cf.Colors[1].Indexed);
                break;
        }

        dataBar.Append(conditionalFormatValueObject1);
        dataBar.Append(conditionalFormatValueObject2);
        dataBar.Append(color);

        conditionalFormattingRule.Append(dataBar);

        ConditionalFormattingRuleExtensionList conditionalFormattingRuleExtensionList = new();
        conditionalFormattingRuleExtensionList.Append(BuildRuleExtension(cf));
        conditionalFormattingRule.Append(conditionalFormattingRuleExtensionList);

        return conditionalFormattingRule;
    }

    private static ConditionalFormattingRuleExtension BuildRuleExtension(IXLConditionalFormat cf)
    {
        ConditionalFormattingRuleExtension conditionalFormattingRuleExtension = new()
        {
            Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
        };
        conditionalFormattingRuleExtension.AddNamespaceDeclaration(
            "x14",
            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
        );
        Id id = new() { Text = ((XLConditionalFormat)cf).Id.WrapInBraces() };
        conditionalFormattingRuleExtension.Append(id);

        return conditionalFormattingRuleExtension;
    }

    private static ConditionalFormatValueObject GetConditionalFormatValueObjectByIndex(
        IXLConditionalFormat cf,
        int index,
        ConditionalFormatValueObjectValues defaultType
    )
    {
        ConditionalFormatValueObject conditionalFormatValueObject = new();

        if (cf.ContentTypes.TryGetValue(index, out XLCFContentType contentType))
        {
            conditionalFormatValueObject.Type = contentType.ToOpenXml();
        }
        else
        {
            conditionalFormatValueObject.Type = defaultType;
        }

        if (cf.Values.TryGetValue(index, out XLFormula? value1) && value1?.Value != null)
        {
            conditionalFormatValueObject.Val = value1.Value;
        }

        return conditionalFormatValueObject;
    }
}

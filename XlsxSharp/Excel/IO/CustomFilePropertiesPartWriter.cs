using System;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using XlsxSharp.Excel.CustomProperties;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.IO;

internal class CustomFilePropertiesPartWriter
{
    internal static void GenerateContent(
        CustomFilePropertiesPart customFilePropertiesPart,
        XLWorkbook workbook
    )
    {
        Properties properties = new();
        properties.AddNamespaceDeclaration(
            "vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        );
        int propertyId = 1;
        foreach (IXLCustomProperty p in workbook.CustomProperties)
        {
            propertyId++;
            CustomDocumentProperty customDocumentProperty = new()
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = propertyId,
                Name = p.Name,
            };
            if (p.Type == XLCustomPropertyType.Text)
            {
                VTLPWSTR vTlpwstr1 = new() { Text = p.GetValue<string>() };
                customDocumentProperty.AppendChild(vTlpwstr1);
            }
            else if (p.Type == XLCustomPropertyType.Date)
            {
                VTFileTime vTFileTime1 = new()
                {
                    Text = p.GetValue<DateTime>()
                        .ToUniversalTime()
                        .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"),
                };
                customDocumentProperty.AppendChild(vTFileTime1);
            }
            else if (p.Type == XLCustomPropertyType.Number)
            {
                VTDouble vTDouble1 = new() { Text = p.GetValue<Double>().ToInvariantString() };
                customDocumentProperty.AppendChild(vTDouble1);
            }
            else
            {
                VTBool vTBool1 = new() { Text = p.GetValue<Boolean>().ToString().ToLower() };
                customDocumentProperty.AppendChild(vTBool1);
            }
            properties.AppendChild(customDocumentProperty);
        }

        customFilePropertiesPart.Properties = properties;
    }
}

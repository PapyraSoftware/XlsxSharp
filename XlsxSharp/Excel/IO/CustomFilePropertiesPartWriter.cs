using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.CustomProperties;
using XlsxSharp.Extensions;
using XlsxSharp.IO.Packaging;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// Writes <c>/docProps/custom.xml</c>. Unlike the extended properties next to it, this part holds
/// nothing but what the workbook model already has, so it is written from scratch every time.
/// </summary>
internal class CustomFilePropertiesPartWriter
{
    private static readonly XNamespace Op =
        "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

    private static readonly XNamespace Vt =
        "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    /// <summary>
    /// The format id of the user defined property set, which is the same constant for every
    /// custom property in every OOXML document (ECMA-376 Part 1 §15.2.12.2).
    /// </summary>
    private const string UserDefinedFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    internal static void GenerateContent(OpcPart customFilePropertiesPart, XLWorkbook workbook)
    {
        XElement properties = new(
            Op + "Properties",
            new XAttribute(XNamespace.Xmlns + "vt", Vt.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "op", Op.NamespaceName)
        );

        // Property ids are 1-based with 1 reserved for the property set itself, so the first
        // custom property is 2.
        int propertyId = 1;
        foreach (IXLCustomProperty property in workbook.CustomProperties)
        {
            propertyId++;
            properties.Add(
                new XElement(
                    Op + "property",
                    new XAttribute("fmtid", UserDefinedFormatId),
                    new XAttribute("pid", propertyId),
                    new XAttribute("name", property.Name),
                    Value(property)
                )
            );
        }

        using Stream partStream = customFilePropertiesPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );

        new XDocument(properties).Save(xml);
    }

    private static XElement Value(IXLCustomProperty property) =>
        property.Type switch
        {
            XLCustomPropertyType.Text => new XElement(Vt + "lpwstr", property.GetValue<string>()),

            // Invariant culture, so that a non-Gregorian calendar culture cannot turn the year
            // into something Excel will not read back.
            XLCustomPropertyType.Date => new XElement(
                Vt + "filetime",
                property
                    .GetValue<DateTime>()
                    .ToUniversalTime()
                    .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'", CultureInfo.InvariantCulture)
            ),

            XLCustomPropertyType.Number => new XElement(
                Vt + "r8",
                property.GetValue<double>().ToInvariantString()
            ),

            _ => new XElement(Vt + "bool", property.GetValue<bool>() ? "true" : "false"),
        };
}

#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.IO;

internal class ExtendedFilePropertiesPartWriter
{
    internal static void GenerateContent(
        ExtendedFilePropertiesPart extendedFilePropertiesPart,
        XLWorkbook workbook
    )
    {
        if (extendedFilePropertiesPart.Properties == null)
        {
            extendedFilePropertiesPart.Properties = new Properties();
        }

        Properties properties = extendedFilePropertiesPart.Properties;
        if (
            !properties.NamespaceDeclarations.Contains(
                new KeyValuePair<string, string>(
                    "vt",
                    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                )
            )
        )
        {
            properties.AddNamespaceDeclaration(
                "vt",
                "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
            );
        }

        if (properties.Application == null)
        {
            properties.AppendChild(new Application { Text = "Microsoft Excel" });
        }

        if (properties.DocumentSecurity == null)
        {
            properties.AppendChild(new DocumentSecurity { Text = "0" });
        }

        if (properties.ScaleCrop == null)
        {
            properties.AppendChild(new ScaleCrop { Text = "false" });
        }

        if (properties.HeadingPairs == null)
        {
            properties.HeadingPairs = new HeadingPairs();
        }

        if (properties.TitlesOfParts == null)
        {
            properties.TitlesOfParts = new TitlesOfParts();
        }

        properties.HeadingPairs.VTVector = new VTVector { BaseType = VectorBaseValues.Variant };

        properties.TitlesOfParts.VTVector = new VTVector { BaseType = VectorBaseValues.Lpstr };

        VTVector vTVectorOne = properties.HeadingPairs.VTVector;

        VTVector vTVectorTwo = properties.TitlesOfParts.VTVector;

        var modifiedWorksheets = ((IEnumerable<XLWorksheet>)workbook.WorksheetsInternal)
            .Select(w => new { w.Name, Order = w.Position })
            .ToList();
        List<string> modifiedNamedRanges = GetModifiedNamedRanges(workbook);
        int modifiedWorksheetsCount = modifiedWorksheets.Count;
        int modifiedNamedRangesCount = modifiedNamedRanges.Count;

        InsertOnVtVector(vTVectorOne, "Worksheets", 0, modifiedWorksheetsCount.ToInvariantString());
        InsertOnVtVector(
            vTVectorOne,
            "Named Ranges",
            2,
            modifiedNamedRangesCount.ToInvariantString()
        );

        vTVectorTwo.Size = (uint)(modifiedNamedRangesCount + modifiedWorksheetsCount);

        foreach (
            VTLPSTR vTlpstr3 in modifiedWorksheets
                .OrderBy(w => w.Order)
                .Select(w => new VTLPSTR { Text = w.Name })
        )
        {
            vTVectorTwo.AppendChild(vTlpstr3);
        }

        foreach (VTLPSTR vTlpstr7 in modifiedNamedRanges.Select(nr => new VTLPSTR { Text = nr }))
        {
            vTVectorTwo.AppendChild(vTlpstr7);
        }

        if (workbook.Properties.Manager != null)
        {
            if (!string.IsNullOrWhiteSpace(workbook.Properties.Manager))
            {
                if (properties.Manager == null)
                {
                    properties.Manager = new Manager();
                }

                properties.Manager.Text = workbook.Properties.Manager;
            }
            else
            {
                properties.Manager = null;
            }
        }

        if (workbook.Properties.Company == null)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(workbook.Properties.Company))
        {
            if (properties.Company == null)
            {
                properties.Company = new Company();
            }

            properties.Company.Text = workbook.Properties.Company;
        }
        else
        {
            properties.Company = null;
        }
    }

    private static void InsertOnVtVector(VTVector vTVector, string property, int index, string text)
    {
        IEnumerable<Variant> m =
            from e1 in vTVector.Elements<Variant>()
            where e1.Elements<VTLPSTR>().Any(e2 => e2.Text == property)
            select e1;
        if (!m.Any())
        {
            if (vTVector.Size == null)
            {
                vTVector.Size = new UInt32Value(0U);
            }

            vTVector.Size += 2U;
            Variant variant1 = new();
            VTLPSTR vTlpstr1 = new() { Text = property };
            variant1.AppendChild(vTlpstr1);
            vTVector.InsertAt(variant1, index);

            Variant variant2 = new();
            VTInt32 vTInt321 = new();
            variant2.AppendChild(vTInt321);
            vTVector.InsertAt(variant2, index + 1);
        }

        int targetIndex = 0;
        foreach (Variant e in vTVector.Elements<Variant>())
        {
            if (e.Elements<VTLPSTR>().Any(e2 => e2.Text == property))
            {
                vTVector.ElementAt(targetIndex + 1).GetFirstChild<VTInt32>().Text = text;
                break;
            }
            targetIndex++;
        }
    }

    private static List<string> GetModifiedNamedRanges(XLWorkbook workbook)
    {
        List<string> namedRanges = [];
        foreach (XLWorksheet sheet in workbook.WorksheetsInternal)
        {
            namedRanges.AddRange(
                sheet.DefinedNames.Select<XLDefinedName, string>(n => sheet.Name + "!" + n.Name)
            );
            namedRanges.Add(sheet.Name + "!Print_Area");
            namedRanges.Add(sheet.Name + "!Print_Titles");
        }
        namedRanges.AddRange(
            workbook.DefinedNamesInternal.Select<XLDefinedName, string>(n => n.Name)
        );
        return namedRanges;
    }
}

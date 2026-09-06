using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.IO.Packaging;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

/// <summary>
/// A writer for table definition part.
/// </summary>
internal class TablePartWriter
{
    internal static void SynchronizeTableParts(
        OpcPackage package,
        XLTables tables,
        OpcPart worksheetPart,
        SaveContext context
    )
    {
        // Remove table definition parts that are not a part of workbook
        foreach (
            OpcPart tableDefinitionPart in worksheetPart.PartsOfType(OoxmlPartTypes.Table).ToList()
        )
        {
            string partId = worksheetPart.Relationships.GetIdOfTarget(tableDefinitionPart.Name);
            bool xlWorkbookContainsTable = tables.Cast<XLTable>().Any(t => t.RelId == partId);
            if (!xlWorkbookContainsTable)
            {
                package.DeletePart(tableDefinitionPart.Name);
            }
        }

        foreach (XLTable xlTable in tables.Cast<XLTable>())
        {
            if (string.IsNullOrEmpty(xlTable.RelId))
            {
                xlTable.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                worksheetPart.AddPartOfType(
                    package,
                    OoxmlPartTypes.Table,
                    relationshipId: xlTable.RelId
                );
            }
        }
    }

    internal static void GenerateTableParts(
        XLTables tables,
        OpcPart worksheetPart,
        SaveContext context
    )
    {
        foreach (XLTable xlTable in tables.Cast<XLTable>())
        {
            string relId = xlTable.RelId;
            OpcPart tableDefinitionPart = worksheetPart.GetRelatedPart(relId);
            GenerateTableDefinitionPartContent(tableDefinitionPart, xlTable, context);
        }
    }

    private static void GenerateTableDefinitionPartContent(
        OpcPart tableDefinitionPart,
        XLTable xlTable,
        SaveContext context
    )
    {
        context.TableId++;
        string tableName = GetTableName(xlTable.Name, context);
        XElement table = new(
            SpreadsheetXml.Main + "table",
            // The parts are written with the prefix the schema's own examples use, which is the
            // one every other part of the package is written with.
            new XAttribute(XNamespace.Xmlns + "x", SpreadsheetXml.Main.NamespaceName),
            new XAttribute("id", context.TableId),
            new XAttribute("name", tableName),
            new XAttribute("displayName", tableName),
            new XAttribute(
                "ref",
                $"{xlTable.RangeAddress.FirstAddress}:{xlTable.RangeAddress.LastAddress}"
            )
        );

        if (!xlTable.ShowHeaderRow)
        {
            table.SetAttributeValue("headerRowCount", 0);
        }

        // A table that shows no totals row says so, rather than saying it has none.
        if (xlTable.ShowTotalsRow)
        {
            table.SetAttributeValue("totalsRowCount", 1);
        }
        else
        {
            table.SetAttributeValue("totalsRowShown", "0");
        }

        if (xlTable.ShowAutoFilter)
        {
            xlTable.AutoFilter.Range = xlTable.ShowTotalsRow
                ? xlTable.Worksheet.Range(
                    xlTable.RangeAddress.FirstAddress.RowNumber,
                    xlTable.RangeAddress.FirstAddress.ColumnNumber,
                    xlTable.RangeAddress.LastAddress.RowNumber - 1,
                    xlTable.RangeAddress.LastAddress.ColumnNumber
                )
                : xlTable.Worksheet.Range(xlTable.RangeAddress);

            XElement autoFilter = new(SpreadsheetXml.Main + "autoFilter");
            WorksheetPartWriter.PopulateAutoFilter(xlTable.AutoFilter, autoFilter);
            table.Add(autoFilter);
        }

        table.Add(WriteTableColumns(xlTable, context));
        table.Add(WriteTableStyleInfo(xlTable));

        using Stream partStream = tableDefinitionPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { CloseOutput = true, Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );
        xml.WriteStartDocument();
        table.WriteTo(xml);
        xml.WriteEndDocument();
    }

    private static XElement WriteTableColumns(XLTable xlTable, SaveContext context)
    {
        XElement tableColumns = new(
            SpreadsheetXml.Main + "tableColumns",
            new XAttribute("count", (uint)xlTable.ColumnCount())
        );

        uint columnId = 0;
        foreach (XLTableField xlField in xlTable.Fields)
        {
            columnId++;
            XElement tableColumn = new(
                SpreadsheetXml.Main + "tableColumn",
                new XAttribute("id", columnId),
                new XAttribute(
                    "name",
                    xlField
                        .Name.Replace("_x000a_", "_x005f_x000a_")
                        .Replace(Environment.NewLine, "_x000a_")
                )
            );

            // OI-29500: the headerRowDxfId attribute shall be absent if the enclosing table element's headerRow
            // attribute value is greater than or equal to 1. The header dxf stores the format while the table
            // doesn't have a header, but user might want to add it back.
            if (xlField.HeaderFormatValue is { } headerDfx)
            {
                tableColumn.SetAttributeValue("headerRowDxfId", context.GetDxfId(headerDfx));
            }

            if (xlField.DataFormatValue is { } dataDfx)
            {
                tableColumn.SetAttributeValue("dataDxfId", context.GetDxfId(dataDfx));
            }

            if (xlField.TotalFormatValue is { } totalsDfx)
            {
                tableColumn.SetAttributeValue("totalsRowDxfId", context.GetDxfId(totalsDfx));
            }

            if (xlField.IsConsistentFormula())
            {
                string formula = xlField
                    .Column.Cells()
                    .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                    .First()
                    .FormulaA1;

                while (formula.StartsWith('=') && formula.Length > 1)
                {
                    formula = formula[1..];
                }

                if (!string.IsNullOrWhiteSpace(formula))
                {
                    tableColumn.Add(
                        new XElement(SpreadsheetXml.Main + "calculatedColumnFormula", formula)
                    );
                }
            }

            if (xlTable.ShowTotalsRow)
            {
                if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                {
                    tableColumn.SetAttributeValue(
                        "totalsRowFunction",
                        xlField.TotalsRowFunction.ToXml()
                    );

                    if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                    {
                        tableColumn.Add(
                            new XElement(
                                SpreadsheetXml.Main + "totalsRowFormula",
                                xlField.TotalsRowFormulaA1
                            )
                        );
                    }
                }

                if (!string.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                {
                    tableColumn.SetAttributeValue("totalsRowLabel", xlField.TotalsRowLabel);
                }
            }

            tableColumns.Add(tableColumn);
        }

        return tableColumns;
    }

    private static XElement WriteTableStyleInfo(XLTable xlTable)
    {
        XElement tableStyleInfo = new(SpreadsheetXml.Main + "tableStyleInfo");
        if (xlTable.Theme != XLTableTheme.None)
        {
            tableStyleInfo.SetAttributeValue("name", xlTable.Theme.Name);
        }

        WorksheetXml.SetBool(tableStyleInfo, "showFirstColumn", xlTable.EmphasizeFirstColumn);
        WorksheetXml.SetBool(tableStyleInfo, "showLastColumn", xlTable.EmphasizeLastColumn);
        WorksheetXml.SetBool(tableStyleInfo, "showRowStripes", xlTable.ShowRowStripes);
        WorksheetXml.SetBool(tableStyleInfo, "showColumnStripes", xlTable.ShowColumnStripes);
        return tableStyleInfo;
    }

    /// <summary>
    /// Reading <c>xl/tables/tableN.xml</c>.
    /// </summary>
    internal static XElement Read(OpcPart part)
    {
        using Stream stream = part.GetReadStream();
        return XDocument.Load(stream).Root
            ?? throw PartStructureException.ExpectedElementNotFound("table");
    }

    private static string GetTableName(string originalTableName, SaveContext context)
    {
        string tableName = originalTableName.RemoveSpecialCharacters();
        string name = tableName;
        if (context.TableNames.Contains(name))
        {
            int i = 1;
            name = tableName + i.ToInvariantString();
            while (context.TableNames.Contains(name))
            {
                i++;
                name = tableName + i.ToInvariantString();
            }
        }

        context.TableNames.Add(name);
        return name;
    }
}

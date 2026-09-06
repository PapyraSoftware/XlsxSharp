#nullable disable

using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Exceptions;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.IO.Packaging;
using XlsxSharp.Utils;
using static XlsxSharp.Excel.IO.OpenXmlConst;
using static XlsxSharp.Excel.XLWorkbook;

namespace XlsxSharp.Excel.IO;

internal class WorksheetPartWriter
{
    internal static void GenerateWorksheetPartContent(
        bool partIsEmpty,
        OpcPackage package,
        OpcPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveOptions options,
        SaveContext context
    )
    {
        if (options.ConsolidateConditionalFormatRanges)
        {
            xlWorksheet.ConditionalFormats.Consolidate();
        }

        XElement worksheet = ReadOrCreateWorksheet(partIsEmpty, worksheetPart);

        // The cells are written from the workbook model rather than from what the part held, so
        // the rows loaded above are dropped before the rest of the sheet is turned into XML.
        worksheet.Element(SpreadsheetXml.Main + "sheetData")?.RemoveNodes();
        WorksheetXml.Child(worksheet, "sheetData");

        WriteTableParts(worksheet, (XLTables)xlWorksheet.Tables);
        WritePictures(worksheet, package, worksheetPart, xlWorksheet, context);
        WriteLegacyDrawing(worksheet, xlWorksheet);
        WriteSheetProperties(worksheet, xlWorksheet);
        WriteDimension(worksheet, xlWorksheet);
        WriteSheetViews(worksheet, xlWorksheet);
        WriteSheetFormatProperties(worksheet, xlWorksheet);
        WriteColumns(worksheet, xlWorksheet, context);
        WriteConditionalFormats(worksheet, xlWorksheet, context);
        WriteDataValidations(worksheet, xlWorksheet, options);
        WriteSparklines(worksheet, xlWorksheet);
        WriteSheetProtection(worksheet, xlWorksheet.Protection);
        WriteAutoFilter(worksheet, xlWorksheet.AutoFilter);
        WriteMergedCells(worksheet, xlWorksheet.Internals.MergedRanges);
        WriteHyperlinks(worksheet, worksheetPart, xlWorksheet, context);
        WritePageSetup(worksheet, xlWorksheet.PageSetup);
        WriteHeaderFooter(worksheet, xlWorksheet.PageSetup);
        WriteBreaks(
            worksheet,
            "rowBreaks",
            xlWorksheet.PageSetup.RowBreaks,
            (uint)xlWorksheet.RangeAddress.LastAddress.RowNumber
        );
        WriteBreaks(
            worksheet,
            "colBreaks",
            xlWorksheet.PageSetup.ColumnBreaks,
            (uint)xlWorksheet.RangeAddress.LastAddress.ColumnNumber
        );

        StreamToPart(worksheet, worksheetPart, xlWorksheet, context, options);
    }

    /// <summary>
    /// The sheet as it was loaded, patched rather than replaced, or a fresh one for a part that
    /// has no content yet.
    /// </summary>
    private static XElement ReadOrCreateWorksheet(bool partIsEmpty, OpcPart worksheetPart)
    {
        XElement worksheet;
        if (!partIsEmpty)
        {
            using Stream stream = worksheetPart.GetReadStream();
            worksheet = ReadWorksheetSkippingSheetData(stream);
        }
        else
        {
            worksheet = new XElement(
                SpreadsheetXml.Main + "worksheet",
                new XAttribute(XNamespace.Xmlns + "x", SpreadsheetXml.Main.NamespaceName)
            );
        }

        // The main namespace is always declared under the prefix "x", the way the SDK's own
        // writer always did regardless of what a loaded file used - a file that declares it as
        // the default namespace, or under some other prefix, is normalised here.
        if (worksheet.GetPrefixOfNamespace(SpreadsheetXml.Main) is not "x")
        {
            worksheet
                .Attributes()
                .Where(attribute =>
                    attribute.IsNamespaceDeclaration
                    && attribute.Value == SpreadsheetXml.Main.NamespaceName
                )
                .ToList()
                .ForEach(attribute => attribute.Remove());
            worksheet.SetAttributeValue(XNamespace.Xmlns + "x", SpreadsheetXml.Main.NamespaceName);
        }

        if (
            worksheet
                .Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration)
                .All(attribute => attribute.Value != RelationshipsNs)
        )
        {
            worksheet.SetAttributeValue(XNamespace.Xmlns + "r", RelationshipsNs);
        }

        // We store the x14ac:dyDescent attribute (if set by a xlRow) in a row element. It's an
        // optional attribute and it needs a declared namespace. To avoid writing namespace to
        // each <x:row> element during streaming, write it to every sheet part ahead of time. The
        // namespace has to be marked as ignorable, because Excel's own validator refuses to
        // validate it otherwise, being an optional extension (see ISO29500 part 3).
        if (
            worksheet
                .Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration)
                .All(attribute => attribute.Value != X14Ac2009SsNs)
        )
        {
            worksheet.SetAttributeValue(XNamespace.Xmlns + "x14ac", X14Ac2009SsNs);
            worksheet.SetAttributeValue(XNamespace.Xmlns + "mc", MarkupCompatibilityNs);
            worksheet.SetAttributeValue(XName.Get("Ignorable", MarkupCompatibilityNs), "x14ac");
        }

        return worksheet;
    }

    /// <summary>
    /// The root element of a loaded worksheet part, with an empty <c>sheetData</c> in place of
    /// whatever rows it held.
    /// </summary>
    /// <remarks>
    /// <see cref="GenerateWorksheetPartContent"/> throws every row away and rebuilds
    /// <c>sheetData</c> wholesale from the workbook model regardless of what the part held, so a
    /// full <see cref="XDocument.Load(Stream)"/> here would materialise a sheet's entire row/cell
    /// tree only to discard it - expensive for a large sheet, and the one part of the document
    /// whose loaded content is never actually read. Everything else about the sheet - dimension,
    /// views, columns, merges, hyperlinks, page setup, an unrecognised extension element, and so
    /// on - is small regardless of row count and is still read in full, through the exact same
    /// <see cref="XNode.ReadFrom(XmlReader)"/> the BCL itself uses to build a <see cref="XDocument"/>
    /// from a reader, so it is preserved exactly as before.
    /// </remarks>
    private static XElement ReadWorksheetSkippingSheetData(Stream stream)
    {
        XmlReaderSettings settings = new()
        {
            IgnoreWhitespace = true,
            DtdProcessing = DtdProcessing.Prohibit,
        };

        using XmlReader reader = XmlReader.Create(stream, settings);
        reader.MoveToContent();
        if (reader.NodeType != XmlNodeType.Element)
        {
            throw PartStructureException.ExpectedElementNotFound("worksheet");
        }

        XElement worksheet = new(XName.Get(reader.LocalName, reader.NamespaceURI));
        CopyAttributes(reader, worksheet);

        bool isEmpty = reader.IsEmptyElement;
        reader.Read();
        if (isEmpty)
        {
            return worksheet;
        }

        while (reader.NodeType != XmlNodeType.EndElement)
        {
            if (
                reader.NodeType == XmlNodeType.Element
                && reader.LocalName == "sheetData"
                && reader.NamespaceURI == SpreadsheetXml.Main.NamespaceName
            )
            {
                XElement sheetData = new(XName.Get(reader.LocalName, reader.NamespaceURI));
                CopyAttributes(reader, sheetData);
                worksheet.Add(sheetData);
                reader.Skip();
            }
            else
            {
                worksheet.Add(XNode.ReadFrom(reader));
            }
        }

        return worksheet;
    }

    /// <summary>
    /// Copies the reader's current element's attributes onto <paramref name="element"/>, leaving
    /// the reader back on the element itself.
    /// </summary>
    /// <remarks>
    /// The namespace URI a namespace-aware <see cref="XmlReader"/> reports for an attribute
    /// already resolves to exactly the <see cref="XName"/> <see cref="XAttribute"/> expects for
    /// an ordinary attribute, prefixed or not, and for a prefixed namespace declaration such as
    /// <c>xmlns:x</c> - <see cref="XName.Get(string, string)"/> with the reported namespace URI
    /// matches <see cref="XNamespace.Xmlns"/> plus the local name in both cases. The one exception
    /// is the bare default-namespace declaration <c>xmlns="..."</c>: the reader reports its
    /// namespace URI as the reserved xmlns namespace too, but <see cref="XAttribute"/> represents
    /// it as the plain unprefixed name "xmlns" with no namespace, and rejects the namespace-
    /// qualified form outright.
    /// </remarks>
    private static void CopyAttributes(XmlReader reader, XElement element)
    {
        if (!reader.MoveToFirstAttribute())
        {
            return;
        }

        do
        {
            XName name =
                reader.Prefix.Length == 0 && reader.LocalName == "xmlns"
                    ? XName.Get("xmlns")
                    : XName.Get(reader.LocalName, reader.NamespaceURI);
            element.Add(new XAttribute(name, reader.Value));
        } while (reader.MoveToNextAttribute());

        reader.MoveToElement();
    }

    private static void WriteCellValue(XmlWriter w, XLCell xlCell, SaveContext context)
    {
        XLDataType dataType = xlCell.DataType;
        if (dataType == XLDataType.Blank)
        {
            return;
        }

        if (dataType == XLDataType.Text)
        {
            string text = xlCell.GetText();
            if (xlCell.HasFormula)
            {
                WriteStringValue(w, text);
            }
            else
            {
                if (xlCell.ShareString)
                {
                    int sharedStringId = context.GetSharedStringId(xlCell, text);
                    w.WriteStartElement("v", Main2006SsNs);
                    w.WriteValue(sharedStringId);
                    w.WriteEndElement();
                }
                else
                {
                    w.WriteStartElement("is", Main2006SsNs);
                    XLImmutableRichText richText = xlCell.RichText;
                    if (richText is not null)
                    {
                        TextSerializer.WriteRichTextElements(w, richText, context);
                    }
                    else
                    {
                        w.WriteStartElement("t", Main2006SsNs);
                        if (text.PreserveSpaces())
                        {
                            w.WritePreserveSpaceAttr();
                        }

                        w.WriteString(text);
                        w.WriteEndElement();
                    }

                    w.WriteEndElement(); // is
                }
            }
        }
        else if (dataType == XLDataType.TimeSpan)
        {
            WriteNumberValue(w, xlCell.Value.GetUnifiedNumber());
        }
        else if (dataType == XLDataType.Number)
        {
            WriteNumberValue(w, xlCell.Value.GetNumber());
        }
        else if (dataType == XLDataType.DateTime)
        {
            // OpenXML SDK validator requires a specific format, in addition to the spec, but can reads many more
            DateTime date = xlCell.GetDateTime();
            if (xlCell.Worksheet.Workbook.Use1904DateSystem)
            {
                date = date.AddDays(-1462);
            }

            WriteNumberValue(w, date.ToSerialDateTime());
        }
        else if (dataType == XLDataType.Boolean)
        {
            WriteStringValue(w, xlCell.GetBoolean() ? TrueValue : FalseValue);
        }
        else if (dataType == XLDataType.Error)
        {
            WriteStringValue(w, xlCell.Value.GetError().ToDisplayString());
        }
        else
        {
            throw new InvalidOperationException();
        }

        static void WriteStringValue(XmlWriter w, string text)
        {
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteString(text);
            w.WriteEndElement();
        }

        static void WriteNumberValue(XmlWriter w, double value)
        {
            w.WriteStartElement("v", Main2006SsNs);
            w.WriteNumberValue(value);
            w.WriteEndElement();
        }
    }

    /// <summary>
    /// <c>autoFilter</c>, which is written only when the sheet has one.
    /// </summary>
    private static void WriteAutoFilter(XElement worksheet, XLAutoFilter xlAutoFilter)
    {
        worksheet.Element(SpreadsheetXml.Main + "autoFilter")?.Remove();
        if (xlAutoFilter.IsEnabled)
        {
            PopulateAutoFilter(xlAutoFilter, WorksheetXml.Child(worksheet, "autoFilter"));
        }
    }

    /// <summary>
    /// The filter, its columns and its sort state. A sheet has one of these and so does every
    /// table, which is why it is filled into an element it is handed.
    /// </summary>
    internal static void PopulateAutoFilter(XLAutoFilter xlAutoFilter, XElement autoFilter)
    {
        IXLRange filterRange = xlAutoFilter.Range;
        autoFilter.SetAttributeValue("ref", filterRange.RangeAddress.ToString());

        foreach ((int columnNumber, XLFilterColumn xlFilterColumn) in xlAutoFilter.Columns)
        {
            XElement filterColumn = new(
                SpreadsheetXml.Main + "filterColumn",
                new XAttribute("colId", (uint)columnNumber - 1)
            );

            filterColumn.Add(
                xlFilterColumn.FilterType switch
                {
                    XLFilterType.Custom => CustomFilters(xlFilterColumn),
                    XLFilterType.TopBottom => TopBottomFilter(xlFilterColumn),
                    XLFilterType.Dynamic => DynamicFilter(xlFilterColumn),
                    XLFilterType.Regular => RegularFilters(xlFilterColumn),
                    _ => throw new NotSupportedException(),
                }
            );

            autoFilter.Add(filterColumn);
        }

        if (xlAutoFilter.Sorted)
        {
            autoFilter.Add(SortState(xlAutoFilter, filterRange));
        }
    }

    private static XElement CustomFilters(XLFilterColumn xlFilterColumn)
    {
        XElement customFilters = new(SpreadsheetXml.Main + "customFilters");
        foreach (XLFilter xlFilter in xlFilterColumn)
        {
            // Since OOXML allows only string, the operand for custom filter must be serialized.
            XElement customFilter = new(
                SpreadsheetXml.Main + "customFilter",
                new XAttribute("val", xlFilter.CustomValue.ToString(CultureInfo.InvariantCulture))
            );

            if (xlFilter.Operator != XLFilterOperator.Equal)
            {
                customFilter.SetAttributeValue("operator", xlFilter.Operator.ToXml());
            }

            if (xlFilter.Connector == XLConnector.And)
            {
                WorksheetXml.SetBool(customFilters, "and", true);
            }

            customFilters.Add(customFilter);
        }

        return customFilters;
    }

    private static XElement TopBottomFilter(XLFilterColumn xlFilterColumn)
    {
        // Although there is a filterVal attribute, populating it seems like more trouble than
        // it's worth due to consistency issues. It's optional, so we can't rely on it during
        // load anyway.
        XElement top10 = new(SpreadsheetXml.Main + "top10");
        WorksheetXml.Set(top10, "val", xlFilterColumn.TopBottomValue);
        WorksheetXml.SetBoolDefault(
            top10,
            "percent",
            xlFilterColumn.TopBottomType == XLTopBottomType.Percent,
            false
        );
        WorksheetXml.SetBoolDefault(
            top10,
            "top",
            xlFilterColumn.TopBottomPart == XLTopBottomPart.Top,
            true
        );
        return top10;
    }

    private static XElement DynamicFilter(XLFilterColumn xlFilterColumn)
    {
        XElement dynamicFilter = new(
            SpreadsheetXml.Main + "dynamicFilter",
            new XAttribute("type", xlFilterColumn.DynamicType.ToXml())
        );
        WorksheetXml.Set(dynamicFilter, "val", xlFilterColumn.DynamicValue);
        return dynamicFilter;
    }

    /// <summary>
    /// The plain value filters, and the date group filters after them - the schema puts every
    /// filter before every dateGroupItem, whatever order the workbook model holds them in.
    /// </summary>
    private static XElement RegularFilters(XLFilterColumn xlFilterColumn)
    {
        XElement filters = new(SpreadsheetXml.Main + "filters");

        foreach (XLFilter filter in xlFilterColumn)
        {
            if (filter.Value is string value)
            {
                filters.Add(
                    new XElement(SpreadsheetXml.Main + "filter", new XAttribute("val", value))
                );
            }
        }

        foreach (XLFilter filter in xlFilterColumn)
        {
            if (filter.Value is DateTime date)
            {
                filters.Add(DateGroupItem(date, filter.DateTimeGrouping));
            }
        }

        return filters;
    }

    /// <summary>
    /// A date named only down to the unit the filter groups by.
    /// </summary>
    private static XElement DateGroupItem(DateTime date, XLDateTimeGrouping grouping)
    {
        XElement dateGroupItem = new(
            SpreadsheetXml.Main + "dateGroupItem",
            new XAttribute("year", (ushort)date.Year),
            new XAttribute("dateTimeGrouping", grouping.ToXml())
        );

        Part(XLDateTimeGrouping.Month, "month", date.Month);
        Part(XLDateTimeGrouping.Day, "day", date.Day);
        Part(XLDateTimeGrouping.Hour, "hour", date.Hour);
        Part(XLDateTimeGrouping.Minute, "minute", date.Minute);
        Part(XLDateTimeGrouping.Second, "second", date.Second);
        return dateGroupItem;

        void Part(XLDateTimeGrouping unit, string name, int value)
        {
            if (grouping >= unit)
            {
                dateGroupItem.SetAttributeValue(name, (ushort)value);
            }
        }
    }

    private static XElement SortState(XLAutoFilter xlAutoFilter, IXLRange filterRange)
    {
        // The sorted range is the filtered one without its header row, unless there is only the
        // one row to sort.
        string reference =
            filterRange.FirstCell().Address.RowNumber < filterRange.LastCell().Address.RowNumber
                ? filterRange
                    .Range(filterRange.FirstCell().CellBelow(), filterRange.LastCell())
                    .RangeAddress.ToString()
                : filterRange.RangeAddress.ToString();

        XElement sortCondition = new(
            SpreadsheetXml.Main + "sortCondition",
            new XAttribute(
                "ref",
                filterRange
                    .Range(
                        1,
                        xlAutoFilter.SortColumn,
                        filterRange.RowCount(),
                        xlAutoFilter.SortColumn
                    )
                    .RangeAddress.ToString()
            )
        );

        if (xlAutoFilter.SortOrder == XLSortOrder.Descending)
        {
            WorksheetXml.SetBool(sortCondition, "descending", true);
        }

        return new XElement(
            SpreadsheetXml.Main + "sortState",
            new XAttribute("ref", reference),
            sortCondition
        );
    }

    /// <summary>
    /// <c>legacyDrawing</c>, the reference to the VML part that carries the sheet's comments.
    /// </summary>
    private static void WriteLegacyDrawing(XElement worksheet, XLWorksheet xlWorksheet)
    {
        worksheet.Elements(SpreadsheetXml.Main + "legacyDrawing").Remove();
        if (!string.IsNullOrEmpty(xlWorksheet.LegacyDrawingId))
        {
            WorksheetXml
                .Child(worksheet, "legacyDrawing")
                .SetAttributeValue(SpreadsheetXml.Rel + "id", xlWorksheet.LegacyDrawingId);
        }
    }

    /// <summary>
    /// The columns, patched onto whatever the sheet already had rather than written fresh - a
    /// loaded workbook can carry column spans this pass never touches (past the last column the
    /// model names explicitly, say), and those are updated in place rather than discarded.
    /// </summary>
    private static void WriteColumns(
        XElement worksheet,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        double worksheetColumnWidth = GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
        uint worksheetStyleId = context.GetStyleId(xlWorksheet.FormatValue);

        if (
            xlWorksheet.Internals.CellsCollection.IsEmpty
            && xlWorksheet.Internals.ColumnsCollection.Count == 0
            && worksheetStyleId == 0
        )
        {
            worksheet.Element(SpreadsheetXml.Main + "cols")?.Remove();
            return;
        }

        XElement columns = WorksheetXml.Child(worksheet, "cols");
        Dictionary<uint, XElement> sheetColumnsByMin = columns
            .Elements(SpreadsheetXml.Main + "col")
            .ToDictionary(col => SpreadsheetXml.UInt(col, "min")!.Value, col => col);

        int minInColumnsCollection;
        int maxInColumnsCollection;
        if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
        {
            minInColumnsCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Min();
            maxInColumnsCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
        }
        else
        {
            minInColumnsCollection = 1;
            maxInColumnsCollection = 0;
        }

        // Columns before the first one the workbook model names explicitly get the sheet's own
        // default, one column at a time - matching the granularity UpdateColumn works at.
        for (int co = 1; co < minInColumnsCollection; co++)
        {
            UpdateColumn(
                NewColumn((uint)co, (uint)co, worksheetStyleId, worksheetColumnWidth),
                columns,
                sheetColumnsByMin
            );
        }

        for (int co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
        {
            uint styleId = worksheetStyleId;
            double columnWidth = worksheetColumnWidth;
            bool isHidden = false;
            bool collapsed = false;
            int outlineLevel = 0;
            if (xlWorksheet.Internals.ColumnsCollection.TryGetValue(co, out XLColumn col))
            {
                styleId = col.FormatValue is null
                    ? worksheetStyleId
                    : context.GetStyleId(col.FormatValue);
                columnWidth = GetColumnWidth(col.Width).SaveRound();
                isHidden = col.IsHidden;
                collapsed = col.Collapsed;
                outlineLevel = col.OutlineLevel;
            }

            UpdateColumn(
                NewColumn(
                    (uint)co,
                    (uint)co,
                    styleId,
                    columnWidth,
                    isHidden ? true : null,
                    collapsed ? true : null,
                    outlineLevel > 0 ? (byte)outlineLevel : null
                ),
                columns,
                sheetColumnsByMin
            );
        }

        // Anything past what the model named explicitly - columns the sheet already carried from
        // being loaded - takes on the sheet's own style and width too.
        int lastExplicitColumn = maxInColumnsCollection;
        foreach (
            XElement col in columns
                .Elements(SpreadsheetXml.Main + "col")
                .Where(col => SpreadsheetXml.UInt(col, "min") > (uint)lastExplicitColumn)
                .OrderBy(col => SpreadsheetXml.UInt(col, "min"))
                .ToList()
        )
        {
            col.SetAttributeValue("style", worksheetStyleId);
            WorksheetXml.Set(col, "width", worksheetColumnWidth);
            col.SetAttributeValue("customWidth", "1");

            uint colMax = SpreadsheetXml.UInt(col, "max")!.Value;
            if (colMax > maxInColumnsCollection)
            {
                maxInColumnsCollection = (int)colMax;
            }
        }

        if (maxInColumnsCollection < XlsxSharp.XLHelper.MaxColumnNumber && worksheetStyleId != 0)
        {
            columns.Add(
                NewColumn(
                    (uint)(maxInColumnsCollection + 1),
                    (uint)XlsxSharp.XLHelper.MaxColumnNumber,
                    worksheetStyleId,
                    worksheetColumnWidth
                )
            );
        }

        CollapseColumns(columns, sheetColumnsByMin);

        if (!columns.Elements(SpreadsheetXml.Main + "col").Any())
        {
            columns.Remove();
        }
    }

    private static double GetColumnWidth(double columnWidth) =>
        Math.Min(255.0, Math.Max(0.0, columnWidth + XLConstants.ColumnWidthOffset));

    private static XElement NewColumn(
        uint min,
        uint max,
        uint style,
        double width,
        bool? hidden = null,
        bool? collapsed = null,
        byte? outlineLevel = null
    )
    {
        XElement column = new(
            SpreadsheetXml.Main + "col",
            new XAttribute("min", min),
            new XAttribute("max", max),
            new XAttribute("style", style),
            new XAttribute("customWidth", "1")
        );
        WorksheetXml.Set(column, "width", width);
        if (hidden == true)
        {
            column.SetAttributeValue("hidden", "1");
        }

        if (collapsed == true)
        {
            column.SetAttributeValue("collapsed", "1");
        }

        if (outlineLevel is { } level && level > 0)
        {
            column.SetAttributeValue("outlineLevel", level);
        }

        return column;
    }

    /// <summary>
    /// Places one column's worth of formatting into the tree, splitting an existing span that
    /// covers it if there is one. A span wider than a single column is shaved down one column at
    /// a time as each of its columns is visited in turn, rather than replaced outright, so a
    /// column the model never mentions keeps whatever of the span's own attributes it had.
    /// </summary>
    private static void UpdateColumn(
        XElement column,
        XElement columns,
        Dictionary<uint, XElement> sheetColumnsByMin
    )
    {
        uint columnMin = SpreadsheetXml.UInt(column, "min")!.Value;
        if (!sheetColumnsByMin.TryGetValue(columnMin, out XElement existingColumn))
        {
            XElement newColumn = new(column);
            columns.Add(newColumn);
            sheetColumnsByMin.Add(columnMin, newColumn);
            return;
        }

        XElement replacement = new(existingColumn);
        replacement.SetAttributeValue("min", columnMin);
        replacement.SetAttributeValue("max", SpreadsheetXml.UInt(column, "max"));
        replacement.SetAttributeValue("style", SpreadsheetXml.UInt(column, "style"));
        WorksheetXml.Set(
            replacement,
            "width",
            SpreadsheetXml.Double(column, "width")!.Value.SaveRound()
        );
        replacement.SetAttributeValue("customWidth", column.Attribute("customWidth")?.Value);
        replacement.SetAttributeValue(
            "hidden",
            column.Attribute("hidden") is not null ? "1" : null
        );
        replacement.SetAttributeValue(
            "collapsed",
            column.Attribute("collapsed") is not null ? "1" : null
        );
        replacement.SetAttributeValue(
            "outlineLevel",
            SpreadsheetXml.UInt(column, "outlineLevel") is { } level && level > 0 ? level : null
        );

        sheetColumnsByMin.Remove(columnMin);
        uint existingMin = SpreadsheetXml.UInt(existingColumn, "min")!.Value;
        uint existingMax = SpreadsheetXml.UInt(existingColumn, "max")!.Value;
        if (existingMin + 1 > existingMax)
        {
            // The existing span was exactly one column wide, so it is fully replaced.
            existingColumn.Remove();
            columns.Add(replacement);
            sheetColumnsByMin.Add(columnMin, replacement);
        }
        else
        {
            // The existing span continues past this column; shrink it from the front instead of
            // removing it, so what is left of it can be found under its new starting column.
            columns.Add(replacement);
            sheetColumnsByMin.Add(columnMin, replacement);
            existingColumn.SetAttributeValue("min", existingMin + 1);
            sheetColumnsByMin.Add(existingMin + 1, existingColumn);
        }
    }

    /// <summary>
    /// Merges adjacent columns that carry the same formatting into a single span, the way Excel
    /// itself writes them.
    /// </summary>
    private static void CollapseColumns(XElement columns, Dictionary<uint, XElement> sheetColumns)
    {
        uint lastMin = 1;
        int count = sheetColumns.Count;
        KeyValuePair<uint, XElement>[] ordered = [.. sheetColumns.OrderBy(entry => entry.Key)];
        for (int i = 0; i < count; i++)
        {
            KeyValuePair<uint, XElement> entry = ordered[i];
            if (i + 1 != count && ColumnsAreEqual(entry.Value, ordered[i + 1].Value))
            {
                continue;
            }

            XElement mergedColumn = new(entry.Value);
            mergedColumn.SetAttributeValue("min", lastMin);
            uint mergedMax = SpreadsheetXml.UInt(mergedColumn, "max")!.Value;

            foreach (
                XElement toRemove in columns
                    .Elements(SpreadsheetXml.Main + "col")
                    .Where(col =>
                        SpreadsheetXml.UInt(col, "min") >= lastMin
                        && SpreadsheetXml.UInt(col, "max") <= mergedMax
                    )
                    .ToList()
            )
            {
                toRemove.Remove();
            }

            columns.Add(mergedColumn);
            lastMin = entry.Key + 1;
        }
    }

    private static bool ColumnsAreEqual(XElement left, XElement right) =>
        NullableEquals(SpreadsheetXml.UInt(left, "style"), SpreadsheetXml.UInt(right, "style"))
        && NullableEquals(
            SpreadsheetXml.Double(left, "width"),
            SpreadsheetXml.Double(right, "width"),
            XlsxSharp.XLHelper.Epsilon
        )
        && NullableEquals(SpreadsheetXml.Bool(left, "hidden"), SpreadsheetXml.Bool(right, "hidden"))
        && NullableEquals(
            SpreadsheetXml.Bool(left, "collapsed"),
            SpreadsheetXml.Bool(right, "collapsed")
        )
        && NullableEquals(
            SpreadsheetXml.UInt(left, "outlineLevel"),
            SpreadsheetXml.UInt(right, "outlineLevel")
        );

    private static bool NullableEquals<T>(T? left, T? right)
        where T : struct, IEquatable<T> =>
        (left is null && right is null)
        || (left is not null && right is not null && left.Value.Equals(right.Value));

    private static bool NullableEquals(double? left, double? right, double epsilon) =>
        (left is null && right is null)
        || (left is not null && right is not null && Math.Abs(left.Value - right.Value) < epsilon);

    // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
    // http://archive.oreilly.com/pub/post/what_is_an_emu.html
    // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
    // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
    // http://archive.oreilly.com/pub/post/what_is_an_emu.html
    // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
    private static long ConvertToEnglishMetricUnits(int pixels, double resolution) =>
        Convert.ToInt64(914400L * pixels / resolution);

    /// <summary>
    /// The drawing part as it was, patched with the picture anchors that changed rather than
    /// rebuilt - a loaded sheet's shapes, text boxes and connectors have no place in the picture
    /// model and are carried through untouched, in the order they were in.
    /// </summary>
    private static (XElement Root, bool Standalone) ReadOrCreateWorksheetDrawing(
        OpcPart drawingsPart
    )
    {
        if (drawingsPart.Length > 0)
        {
            using Stream stream = drawingsPart.GetReadStream();
            XDocument existing = XDocument.Load(stream);
            XElement root =
                existing.Root ?? throw PartStructureException.ExpectedElementNotFound("wsDr");
            bool standalone = string.Equals(
                existing.Declaration?.Standalone,
                "yes",
                StringComparison.OrdinalIgnoreCase
            );
            return (root, standalone);
        }

        XElement fresh = new(
            DrawingXml.Xdr + "wsDr",
            new XAttribute(XNamespace.Xmlns + "xdr", DrawingXml.Xdr.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "a", DrawingXml.A.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", SpreadsheetXml.Rel.NamespaceName)
        );
        return (fresh, false);
    }

    private static void SaveWorksheetDrawing(
        OpcPart drawingsPart,
        XElement worksheetDrawing,
        bool standalone
    )
    {
        using Stream stream = drawingsPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            stream,
            new XmlWriterSettings { CloseOutput = true, Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );
        XDocument document = standalone
            ? new XDocument(new XDeclaration("1.0", "utf-8", "yes"), worksheetDrawing)
            : new XDocument(worksheetDrawing);
        document.Save(xml);
    }

    /// <summary>
    /// The content type and file extension of a picture's format. Every format but the two
    /// XlsxSharp added itself (Unknown, Webp) mirrors what the SDK declared for the equivalent
    /// <c>ImagePartType</c>.
    /// </summary>
    private static readonly Dictionary<
        XLPictureFormat,
        (string ContentType, string Extension)
    > ImageContentTypes = new()
    {
        [XLPictureFormat.Unknown] = ("image/unknown", ".bin"),
        [XLPictureFormat.Bmp] = ("image/bmp", ".bmp"),
        [XLPictureFormat.Gif] = ("image/gif", ".gif"),
        [XLPictureFormat.Png] = ("image/png", ".png"),
        [XLPictureFormat.Tiff] = ("image/tiff", ".tiff"),
        [XLPictureFormat.Icon] = ("image/x-icon", ".ico"),
        [XLPictureFormat.Pcx] = ("image/x-pcx", ".pcx"),
        [XLPictureFormat.Jpeg] = ("image/jpeg", ".jpg"),
        [XLPictureFormat.Emf] = ("image/x-emf", ".emf"),
        [XLPictureFormat.Wmf] = ("image/x-wmf", ".wmf"),
        [XLPictureFormat.Webp] = ("image/webp", ".webp"),
    };

    /// <summary>
    /// The first image part name of that extension no part uses yet. The SDK counts separately
    /// per extension and, unlike every other numbered part kind, leaves the first of each
    /// extension unnumbered.
    /// </summary>
    private static string NextFreeImagePartName(OpcPackage package, string extension)
    {
        string first = $"/xl/media/image{extension}";
        if (!package.TryGetPart(first, out _))
        {
            return first;
        }

        for (int number = 2; ; number++)
        {
            string candidate = $"/xl/media/image{number}{extension}";
            if (!package.TryGetPart(candidate, out _))
            {
                return candidate;
            }
        }
    }

    private static void AddPictureAnchor(
        XElement worksheetDrawing,
        OpcPackage package,
        OpcPart drawingsPart,
        Drawings.IXLPicture picture,
        SaveContext context
    )
    {
        XLPicture pic = picture as Drawings.XLPicture;

        // Overwrite actual image binary data
        OpcPart imagePart;
        if (
            !string.IsNullOrEmpty(pic.RelId)
            && drawingsPart.Relationships.TryGetById(pic.RelId, out _)
        )
        {
            imagePart = drawingsPart.GetRelatedPart(pic.RelId);
        }
        else
        {
            pic.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            (string contentType, string extension) = ImageContentTypes[pic.Format];
            (imagePart, _) = drawingsPart.AddPartOfType(
                package,
                OoxmlPartTypes.Image,
                contentType: contentType,
                partName: NextFreeImagePartName(package, extension),
                relationshipId: pic.RelId
            );
        }

        using (MemoryStream stream = new())
        {
            pic.ImageStream.Position = 0;
            pic.ImageStream.CopyTo(stream);
            stream.Seek(0, SeekOrigin.Begin);
            using Stream writeStream = imagePart.GetWriteStream();
            stream.CopyTo(writeStream);
        }

        string embedId = drawingsPart.Relationships.GetIdOfTarget(imagePart.Name);

        // Find the anchor this picture already had, if it has one, so it can be replaced in
        // place rather than moved to the end.
        XElement existingAnchor = worksheetDrawing
            .Elements()
            .FirstOrDefault(anchor => DrawingXml.PictureRelId(anchor) == pic.RelId);

        XLWorkbook wb = pic.Worksheet.Workbook;
        long extentsCx = ConvertToEnglishMetricUnits(pic.Width, wb.DpiX);
        long extentsCy = ConvertToEnglishMetricUnits(pic.Height, wb.DpiY);
        uint nvpId = NextNonVisualDrawingPropertiesId(worksheetDrawing);

        XElement anchor = pic.Placement switch
        {
            Drawings.XLPicturePlacement.FreeFloating => AbsoluteAnchor(
                pic,
                wb,
                extentsCx,
                extentsCy,
                nvpId,
                embedId
            ),
            Drawings.XLPicturePlacement.MoveAndSize => TwoCellAnchor(
                pic,
                wb,
                extentsCx,
                extentsCy,
                nvpId,
                embedId
            ),
            Drawings.XLPicturePlacement.Move => OneCellAnchor(
                pic,
                wb,
                extentsCx,
                extentsCy,
                nvpId,
                embedId
            ),
            _ => null,
        };

        if (anchor is null)
        {
            return;
        }

        if (existingAnchor is not null)
        {
            existingAnchor.ReplaceWith(anchor);
        }
        else
        {
            worksheetDrawing.Add(anchor);
        }
    }

    private static XElement AbsoluteAnchor(
        XLPicture pic,
        XLWorkbook wb,
        long extentsCx,
        long extentsCy,
        uint nvpId,
        string embedId
    ) =>
        new(
            DrawingXml.Xdr + "absoluteAnchor",
            new XElement(
                DrawingXml.Xdr + "pos",
                new XAttribute("x", ConvertToEnglishMetricUnits(pic.Left, wb.DpiX)),
                new XAttribute("y", ConvertToEnglishMetricUnits(pic.Top, wb.DpiY))
            ),
            Extent(extentsCx, extentsCy),
            PictureElement(pic, nvpId, extentsCx, extentsCy, embedId),
            new XElement(DrawingXml.Xdr + "clientData")
        );

    private static XElement TwoCellAnchor(
        XLPicture pic,
        XLWorkbook wb,
        long extentsCx,
        long extentsCy,
        uint nvpId,
        string embedId
    )
    {
        XLMarker from =
            pic.Markers[Drawings.XLMarkerPosition.TopLeft]
            ?? new Drawings.XLMarker(pic.Worksheet.Cell("A1"));
        XLMarker to =
            pic.Markers[Drawings.XLMarkerPosition.BottomRight]
            ?? new Drawings.XLMarker(
                pic.Worksheet.Cell("A1"),
                new System.Drawing.Point(pic.Width, pic.Height)
            );

        return new XElement(
            DrawingXml.Xdr + "twoCellAnchor",
            Marker("from", from, wb),
            Marker("to", to, wb),
            PictureElement(pic, nvpId, extentsCx, extentsCy, embedId),
            new XElement(DrawingXml.Xdr + "clientData")
        );
    }

    private static XElement OneCellAnchor(
        XLPicture pic,
        XLWorkbook wb,
        long extentsCx,
        long extentsCy,
        uint nvpId,
        string embedId
    )
    {
        XLMarker from =
            pic.Markers[Drawings.XLMarkerPosition.TopLeft]
            ?? new Drawings.XLMarker(pic.Worksheet.Cell("A1"));

        return new XElement(
            DrawingXml.Xdr + "oneCellAnchor",
            Marker("from", from, wb),
            Extent(extentsCx, extentsCy),
            PictureElement(pic, nvpId, extentsCx, extentsCy, embedId),
            new XElement(DrawingXml.Xdr + "clientData")
        );
    }

    /// <summary>
    /// A <c>from</c> or <c>to</c> marker: a cell and pixel offset written as EMU child elements
    /// rather than attributes.
    /// </summary>
    private static XElement Marker(string name, XLMarker marker, XLWorkbook wb) =>
        new(
            DrawingXml.Xdr + name,
            new XElement(DrawingXml.Xdr + "col", marker.ColumnNumber - 1),
            new XElement(
                DrawingXml.Xdr + "colOff",
                ConvertToEnglishMetricUnits(marker.Offset.X, wb.DpiX)
            ),
            new XElement(DrawingXml.Xdr + "row", marker.RowNumber - 1),
            new XElement(
                DrawingXml.Xdr + "rowOff",
                ConvertToEnglishMetricUnits(marker.Offset.Y, wb.DpiY)
            )
        );

    private static XElement Extent(long cx, long cy) =>
        new(DrawingXml.Xdr + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy));

    /// <summary>
    /// The extents of a shape's own transform, in the drawingml namespace rather than the
    /// spreadsheet drawing one the anchor-level extent uses.
    /// </summary>
    private static XElement TransformExtent(long cx, long cy) =>
        new(DrawingXml.A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy));

    private static XElement PictureElement(
        XLPicture pic,
        uint nvpId,
        long extentsCx,
        long extentsCy,
        string embedId
    ) =>
        new(
            DrawingXml.Xdr + "pic",
            new XElement(
                DrawingXml.Xdr + "nvPicPr",
                new XElement(
                    DrawingXml.Xdr + "cNvPr",
                    new XAttribute("id", nvpId),
                    new XAttribute("name", pic.Name)
                ),
                new XElement(
                    DrawingXml.Xdr + "cNvPicPr",
                    new XElement(DrawingXml.A + "picLocks", new XAttribute("noChangeAspect", "1"))
                )
            ),
            new XElement(
                DrawingXml.Xdr + "blipFill",
                new XElement(
                    DrawingXml.A + "blip",
                    new XAttribute(SpreadsheetXml.Rel + "embed", embedId),
                    new XAttribute("cstate", "print")
                ),
                new XElement(DrawingXml.A + "stretch", new XElement(DrawingXml.A + "fillRect"))
            ),
            new XElement(
                DrawingXml.Xdr + "spPr",
                new XElement(
                    DrawingXml.A + "xfrm",
                    new XElement(
                        DrawingXml.A + "off",
                        new XAttribute("x", 0),
                        new XAttribute("y", 0)
                    ),
                    TransformExtent(extentsCx, extentsCy)
                ),
                new XElement(DrawingXml.A + "prstGeom", new XAttribute("prst", "rect"))
            )
        );

    /// <summary>
    /// Makes sure the root declares every namespace the drawing uses - which is where the SDK
    /// always put them in addition to wherever else they were declared, regardless of where a
    /// loaded part itself declared them. A shape written by Excel typically declares an
    /// extension's namespace locally, on the element that first needs it, and a newly built
    /// picture anchor has nowhere of its own to declare "r" at all; both are topped up here,
    /// under whatever prefix is already in use, without touching a declaration that already
    /// exists somewhere in the tree.
    /// </summary>
    private static void HoistNamespaceDeclarations(XElement root)
    {
        // Seeded with the three namespaces a freshly built anchor introduces, since a picture
        // anchor being replaced can take its own local declaration of one of these down with it
        // - the relationships namespace in particular is often declared only on the blip that
        // uses it, never on the root, in a file Excel wrote.
        Dictionary<XNamespace, string> prefixes = new()
        {
            [DrawingXml.Xdr] = "xdr",
            [DrawingXml.A] = "a",
            [SpreadsheetXml.Rel] = "r",
        };
        foreach (XElement element in root.DescendantsAndSelf())
        {
            Record(element.Name);
            foreach (XAttribute attribute in element.Attributes())
            {
                if (!attribute.IsNamespaceDeclaration)
                {
                    Record(attribute.Name);
                }
            }

            void Record(XName name)
            {
                if (name.Namespace != XNamespace.None && !prefixes.ContainsKey(name.Namespace))
                {
                    prefixes[name.Namespace] =
                        element.GetPrefixOfNamespace(name.Namespace)
                        ?? throw new InvalidOperationException(
                            $"No prefix in scope for namespace '{name.NamespaceName}'."
                        );
                }
            }
        }

        foreach ((XNamespace ns, string prefix) in prefixes)
        {
            if (root.GetPrefixOfNamespace(ns) is null)
            {
                root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
            }
        }
    }

    /// <summary>
    /// One more than the largest id any shape, connector or picture in the drawing already
    /// carries - ids are shared across every kind of anchor, not just pictures.
    /// </summary>
    private static uint NextNonVisualDrawingPropertiesId(XElement worksheetDrawing)
    {
        List<uint> ids =
        [
            .. worksheetDrawing
                .Descendants(DrawingXml.Xdr + "cNvPr")
                .Select(el => SpreadsheetXml.UInt(el, "id")!.Value),
        ];
        return ids.Count == 0 ? 1U : ids.Max() + 1;
    }

    /// <summary>
    /// Ids are shared across every anchor in the drawing, picture or not, so they are renumbered
    /// as a whole in document order whenever a picture is added or changed.
    /// </summary>
    private static void RebaseNonVisualDrawingPropertiesIds(XElement worksheetDrawing)
    {
        List<XElement> toRebase = [.. worksheetDrawing.Descendants(DrawingXml.Xdr + "cNvPr")];
        for (int i = 0; i < toRebase.Count; i++)
        {
            toRebase[i].SetAttributeValue("id", i + 1);
        }
    }

    /// <summary>
    /// <c>tableParts</c>, which the sheet always carries once any table has ever been added -
    /// an empty tableParts is written rather than removed, matching what the loaded sheet had.
    /// </summary>
    private static void WriteTableParts(XElement worksheet, XLTables xlTables)
    {
        XLTable emptyTable = xlTables.FirstOrDefault<XLTable>(t => t.DataRange is null);
        if (emptyTable != null)
        {
            throw new EmptyTableException($"Table '{emptyTable.Name}' should have at least 1 row.");
        }

        XElement tableParts = WorksheetXml.Child(worksheet, "tableParts");

        xlTables.Deleted.Clear();
        tableParts.RemoveNodes();
        foreach (XLTable xlTable in xlTables.Cast<XLTable>())
        {
            tableParts.Add(
                new XElement(
                    SpreadsheetXml.Main + "tablePart",
                    new XAttribute(SpreadsheetXml.Rel + "id", xlTable.RelId)
                )
            );
        }

        WorksheetXml.Set(tableParts, "count", (uint)xlTables.Count<XLTable>());
    }

    /// <summary>
    /// Everything to do with pictures that isn't the anchors themselves: the parts of deleted
    /// pictures, the sheet's own reference to the drawing part, and dropping that part and the
    /// reference to it once nothing needs it any more.
    /// </summary>
    private static void WritePictures(
        XElement worksheet,
        OpcPackage package,
        OpcPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        OpcPart existingDrawingsPart = worksheetPart.PartOfType(OoxmlPartTypes.Drawing);
        if (existingDrawingsPart is not null)
        {
            XLPictures xlPictures = xlWorksheet.Pictures as Drawings.XLPictures;
            foreach (string removedPicture in xlPictures.Deleted)
            {
                if (existingDrawingsPart.GetRelatedPartOrDefault(removedPicture) is { } imagePart)
                {
                    package.DeletePart(imagePart.Name);
                }
            }
            xlPictures.Deleted.Clear();
        }

        if (xlWorksheet.Pictures.Count > 0)
        {
            OpcPart drawingsPart =
                existingDrawingsPart
                ?? worksheetPart
                    .AddPartOfType(
                        package,
                        OoxmlPartTypes.Drawing,
                        relationshipId: context.RelIdGenerator.GetNext(RelType.Workbook)
                    )
                    .Part;
            (XElement worksheetDrawingXml, bool standalone) = ReadOrCreateWorksheetDrawing(
                drawingsPart
            );

            foreach (XLPicture pic in xlWorksheet.Pictures)
            {
                AddPictureAnchor(worksheetDrawingXml, package, drawingsPart, pic, context);
            }

            RebaseNonVisualDrawingPropertiesIds(worksheetDrawingXml);
            HoistNamespaceDeclarations(worksheetDrawingXml);
            SaveWorksheetDrawing(drawingsPart, worksheetDrawingXml, standalone);

            // A sheet that already carries a drawing reference keeps it exactly as it was
            // loaded; only a sheet gaining pictures for the first time gets one written.
            if (worksheet.Element(SpreadsheetXml.Main + "drawing") is null)
            {
                XElement drawingElement = WorksheetXml.Child(worksheet, "drawing");
                // The SDK always redeclared "r" locally on a newly created element too,
                // redundantly with the root's own declaration, and the reference workbooks
                // record that redundancy.
                drawingElement.SetAttributeValue(XNamespace.Xmlns + "r", RelationshipsNs);
                drawingElement.SetAttributeValue(
                    SpreadsheetXml.Rel + "id",
                    worksheetPart.Relationships.GetIdOfTarget(drawingsPart.Name)
                );
            }
        }

        // Instead of saving a file with an empty Drawings.xml file, rather remove the .xml file
        OpcPart drawingsPartNow = worksheetPart.PartOfType(OoxmlPartTypes.Drawing);
        bool hasCharts =
            drawingsPartNow is not null
            && drawingsPartNow.Relationships.Any(r => r.TargetMode == OpcTargetMode.Internal);
        if (
            drawingsPartNow is not null
            && // There is a drawing part for the sheet that could be deleted
            xlWorksheet.LegacyDrawingId is null
            && // and sheet doesn't contain any form controls or comments or other shapes
            xlWorksheet.Pictures.Count == 0
            && // and also no pictures.
            !hasCharts
        ) // and no charts
        {
            worksheet.Element(SpreadsheetXml.Main + "drawing")?.Remove();
            package.DeletePart(drawingsPartNow.Name);
        }
    }

    /// <summary>
    /// Stream detached worksheet DOM to the worksheet part stream.
    /// Replaces the content of the part.
    /// </summary>
    /// <summary>
    /// <c>sheetPr</c>, which carries the tab colour and how the sheet's outlines are laid out.
    /// </summary>
    private static void WriteSheetProperties(XElement worksheet, XLWorksheet xlWorksheet)
    {
        XElement sheetProperties = WorksheetXml.Child(worksheet, "sheetPr");

        sheetProperties.Element(SpreadsheetXml.Main + "tabColor")?.Remove();
        if (xlWorksheet.TabColor.HasValue)
        {
            SpreadsheetXml.SetColor(
                WorksheetXml.Child(sheetProperties, "tabColor", WorksheetXml.SheetPropertyOrder),
                xlWorksheet.TabColor
            );
        }

        XElement outline = WorksheetXml.Child(
            sheetProperties,
            "outlinePr",
            WorksheetXml.SheetPropertyOrder
        );
        WorksheetXml.SetBool(
            outline,
            "summaryBelow",
            xlWorksheet.Outline.SummaryVLocation == XLOutlineSummaryVLocation.Bottom
        );
        WorksheetXml.SetBool(
            outline,
            "summaryRight",
            xlWorksheet.Outline.SummaryHLocation == XLOutlineSummaryHLocation.Right
        );

        // A sheet set to fit to a number of pages says so here; the counts themselves are on
        // pageSetup. A sheet that already says it is left alone.
        if (
            sheetProperties.Element(SpreadsheetXml.Main + "pageSetUpPr") is null
            && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0)
        )
        {
            WorksheetXml.SetBool(
                WorksheetXml.Child(sheetProperties, "pageSetUpPr", WorksheetXml.SheetPropertyOrder),
                "fitToPage",
                true
            );
        }
    }

    /// <summary>
    /// <c>dimension</c>, which is only ever set once - a sheet that already has one from being
    /// loaded keeps whatever it said, stale or not.
    /// </summary>
    private static void WriteDimension(XElement worksheet, XLWorksheet xlWorksheet)
    {
        if (worksheet.Element(SpreadsheetXml.Main + "dimension") is not null)
        {
            return;
        }

        // Empty worksheets have dimension A1 (not A1:A1)
        string reference = "A1";
        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            int maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
            int maxRow = xlWorksheet.Internals.CellsCollection.MaxRowUsed;
            reference =
                "A1:"
                + XlsxSharp.XLHelper.GetColumnLetterFromNumber(maxColumn)
                + maxRow.ToInvariantString();
        }

        WorksheetXml.Child(worksheet, "dimension").SetAttributeValue("ref", reference);
    }

    /// <summary>
    /// <c>sheetViews</c>, and within it the one <c>sheetView</c> the workbook model tracks.
    /// </summary>
    private static void WriteSheetViews(XElement worksheet, XLWorksheet xlWorksheet)
    {
        XElement sheetViews = WorksheetXml.Child(worksheet, "sheetViews");
        XElement sheetView = sheetViews.Element(SpreadsheetXml.Main + "sheetView");
        if (sheetView is null)
        {
            sheetView = new XElement(
                SpreadsheetXml.Main + "sheetView",
                new XAttribute("workbookViewId", 0)
            );
            sheetViews.Add(sheetView);
        }

        WorksheetXml.SetBoolOptional(
            sheetView,
            "tabSelected",
            xlWorksheet.TabSelected ? true : null
        );
        WorksheetXml.SetBoolOptional(
            sheetView,
            "rightToLeft",
            xlWorksheet.RightToLeft ? true : null
        );
        WorksheetXml.SetBoolOptional(
            sheetView,
            "showFormulas",
            xlWorksheet.ShowFormulas ? true : null
        );

        // These five default to shown; only an explicit "0" turns them off.
        HideWhenFalse(sheetView, "showGridLines", xlWorksheet.ShowGridLines);
        HideWhenFalse(sheetView, "showOutlineSymbols", xlWorksheet.ShowOutlineSymbols);
        HideWhenFalse(sheetView, "showRowColHeaders", xlWorksheet.ShowRowColHeaders);
        HideWhenFalse(sheetView, "showRuler", xlWorksheet.ShowRuler);
        HideWhenFalse(sheetView, "showWhiteSpace", xlWorksheet.ShowWhiteSpace);
        HideWhenFalse(sheetView, "showZeros", xlWorksheet.ShowZeros);

        sheetView.SetAttributeValue(
            "view",
            xlWorksheet.SheetView.View == XLSheetViewOptions.Normal
                ? null
                : xlWorksheet.SheetView.View.ToXml()
        );

        XElement pane = WritePane(sheetView, xlWorksheet, out int hSplit, out int ySplit);

        // Whether it's for a regular sheet or the bottom-right pane, the top left cell of the
        // view is only written when it differs from the sheet's own default.
        sheetView.SetAttributeValue(
            "topLeftCell",
            !xlWorksheet.SheetView.TopLeftCellAddress.IsValid
            || xlWorksheet.SheetView.TopLeftCellAddress
                == new XLAddress(1, 1, fixedRow: false, fixedColumn: false)
                ? null
                : xlWorksheet.SheetView.TopLeftCellAddress.ToString()
        );

        WriteSelections(sheetView, xlWorksheet, pane);

        WriteZoom(sheetView, "zoomScale", xlWorksheet.SheetView.ZoomScale);
        WriteZoom(sheetView, "zoomScaleNormal", xlWorksheet.SheetView.ZoomScaleNormal);
        WriteZoom(
            sheetView,
            "zoomScalePageLayoutView",
            xlWorksheet.SheetView.ZoomScalePageLayoutView
        );
        WriteZoom(
            sheetView,
            "zoomScaleSheetLayoutView",
            xlWorksheet.SheetView.ZoomScaleSheetLayoutView
        );

        static void HideWhenFalse(XElement element, string name, bool shown) =>
            element.SetAttributeValue(name, shown ? null : "0");

        static void WriteZoom(XElement element, string name, int zoom) =>
            element.SetAttributeValue(
                name,
                zoom == 100 ? null : (uint)Math.Max(10, Math.Min(400, zoom))
            );
    }

    /// <summary>
    /// The frozen pane, if the sheet has a split. Only <see cref="XLSheetViewOptions"/> that split
    /// the sheet ever get written - a plain scroll split has nowhere to go in the workbook model.
    /// </summary>
    private static XElement WritePane(
        XElement sheetView,
        XLWorksheet xlWorksheet,
        out int hSplit,
        out int ySplit
    )
    {
        hSplit = xlWorksheet.SheetView.SplitColumn;
        ySplit = xlWorksheet.SheetView.SplitRow;

        if (hSplit == 0 && ySplit == 0)
        {
            sheetView.Elements(SpreadsheetXml.Main + "pane").Remove();
            return null;
        }

        XElement pane =
            sheetView.Element(SpreadsheetXml.Main + "pane")
            ?? WorksheetXml.Child(sheetView, "pane", WorksheetXml.SheetViewOrder);

        pane.RemoveAttributes();
        pane.SetAttributeValue("state", "frozenSplit");
        pane.SetAttributeValue("xSplit", hSplit);
        pane.SetAttributeValue("ySplit", ySplit);

        // When panes are frozen, which part should move.
        string activePane = (ySplit: ySplit != 0, hSplit: hSplit != 0) switch
        {
            (false, false) => "topLeft",
            (false, true) => "topRight",
            (true, false) => "bottomLeft",
            (true, true) => "bottomRight",
        };
        pane.SetAttributeValue("activePane", activePane);
        pane.SetAttributeValue(
            "topLeftCell",
            XlsxSharp.XLHelper.GetColumnLetterFromNumber(hSplit + 1) + (ySplit + 1)
        );

        return pane;
    }

    private static void WriteSelections(XElement sheetView, XLWorksheet xlWorksheet, XElement pane)
    {
        if (!xlWorksheet.SelectedRanges.Any() && xlWorksheet.ActiveCell is null)
        {
            return;
        }

        sheetView.Elements(SpreadsheetXml.Main + "selection").Remove();

        IXLRange firstSelection = xlWorksheet.SelectedRanges.FirstOrDefault();

        // If a pane exists, we need to set the active pane too. Yes, this might lead to 2
        // Selection elements!
        if (pane is not null)
        {
            AddSelection(pane.Attribute("activePane")?.Value);
        }

        AddSelection(null);

        void AddSelection(string activePane)
        {
            XElement selection = new(SpreadsheetXml.Main + "selection");
            if (activePane is not null)
            {
                selection.SetAttributeValue("pane", activePane);
            }

            string activeCell = xlWorksheet.ActiveCell is not null
                ? xlWorksheet.ActiveCell.Value.ToString()
                : firstSelection?.RangeAddress.FirstAddress.ToStringRelative(false);
            selection.SetAttributeValue("activeCell", activeCell);

            List<string> sequence =
            [
                activeCell,
                .. xlWorksheet.SelectedRanges.Select(range =>
                    range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress)
                        ? range.RangeAddress.FirstAddress.ToStringRelative(false)
                        : range.RangeAddress.ToStringRelative(false)
                ),
            ];
            selection.SetAttributeValue("sqref", string.Join(" ", sequence.Distinct()));

            WorksheetXml.Insert(sheetView, "selection", selection, WorksheetXml.SheetViewOrder);
        }
    }

    /// <summary>
    /// <c>sheetFormatPr</c>, which carries the sheet's default row height and column width and
    /// how deep its outlines go.
    /// </summary>
    private static void WriteSheetFormatProperties(XElement worksheet, XLWorksheet xlWorksheet)
    {
        XElement element = WorksheetXml.Child(worksheet, "sheetFormatPr");

        WorksheetXml.Set(element, "defaultRowHeight", xlWorksheet.RowHeight.SaveRound());
        WorksheetXml.SetBoolOptional(
            element,
            "customHeight",
            xlWorksheet.RowHeightChanged ? true : null
        );
        WorksheetXml.SetOptional<double>(
            element,
            "defaultColWidth",
            xlWorksheet.ColumnWidthChanged
                ? GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound()
                : null
        );

        int maxOutlineColumn =
            xlWorksheet.ColumnCount() > 0 ? xlWorksheet.GetMaxColumnOutline() : 0;
        int maxOutlineRow = xlWorksheet.RowCount() > 0 ? xlWorksheet.GetMaxRowOutline() : 0;
        WorksheetXml.SetOptional<byte>(
            element,
            "outlineLevelCol",
            maxOutlineColumn > 0 ? (byte)maxOutlineColumn : null
        );
        WorksheetXml.SetOptional<byte>(
            element,
            "outlineLevelRow",
            maxOutlineRow > 0 ? (byte)maxOutlineRow : null
        );
    }

    /// <summary>
    /// The sparkline groups, which live in an x14 extension of their own - the 2006 schema
    /// predates them.
    /// </summary>
    private static void WriteSparklines(XElement worksheet, XLWorksheet xlWorksheet)
    {
        const string uri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

        if (!xlWorksheet.SparklineGroups.Any())
        {
            RemoveExtension(worksheet, uri, SpreadsheetXml.X14 + "sparklineGroups");
            return;
        }

        XElement sparklineGroups = Extension(worksheet, uri, "sparklineGroups");

        foreach (XLSparklineGroup xlSparklineGroup in xlWorksheet.SparklineGroupsInternal)
        {
            // Do not create an empty Sparkline group
            if (!xlSparklineGroup.Sparklines.Any())
            {
                continue;
            }

            sparklineGroups.Add(WriteSparklineGroup(xlSparklineGroup));
        }

        // if all Sparkline groups had no Sparklines, remove the entire SparklineGroup element
        if (!sparklineGroups.Elements().Any())
        {
            sparklineGroups.Remove();
        }
    }

    private static XElement WriteSparklineGroup(XLSparklineGroup xlSparklineGroup)
    {
        XNamespace revision2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
        XElement sparklineGroup = new(
            SpreadsheetXml.X14 + "sparklineGroup",
            new XAttribute(XNamespace.Xmlns + "xr2", revision2.NamespaceName),
            new XAttribute(revision2 + "uid", "{A98FF5F8-AE60-43B5-8001-AD89004F45D3}")
        );

        WorksheetXml.Set(sparklineGroup, "lineWeight", xlSparklineGroup.LineWeight);
        sparklineGroup.SetAttributeValue("type", xlSparklineGroup.Type.ToXml());
        sparklineGroup.SetAttributeValue(
            "displayEmptyCellsAs",
            xlSparklineGroup.DisplayEmptyCellsAs.ToXml()
        );
        WorksheetXml.SetBool(sparklineGroup, "displayHidden", xlSparklineGroup.DisplayHidden);

        Marker("markers", XLSparklineMarkers.Markers);
        Marker("high", XLSparklineMarkers.HighPoint);
        Marker("low", XLSparklineMarkers.LowPoint);
        Marker("first", XLSparklineMarkers.FirstPoint);
        Marker("last", XLSparklineMarkers.LastPoint);
        Marker("negative", XLSparklineMarkers.NegativePoints);

        IXLSparklineHorizontalAxis horizontalAxis = xlSparklineGroup.HorizontalAxis;
        WorksheetXml.SetBool(sparklineGroup, "displayXAxis", horizontalAxis.IsVisible);
        WorksheetXml.SetBool(sparklineGroup, "rightToLeft", horizontalAxis.RightToLeft);
        WorksheetXml.SetBool(sparklineGroup, "dateAxis", horizontalAxis.DateAxis);

        IXLSparklineVerticalAxis verticalAxis = xlSparklineGroup.VerticalAxis;
        sparklineGroup.SetAttributeValue("minAxisType", verticalAxis.MinAxisType.ToXml());
        sparklineGroup.SetAttributeValue("maxAxisType", verticalAxis.MaxAxisType.ToXml());

        // A bound is only named when the axis is set to a bound of its own.
        if (verticalAxis.MinAxisType == XLSparklineAxisMinMax.Custom)
        {
            WorksheetXml.SetOptional(sparklineGroup, "manualMin", verticalAxis.ManualMin);
        }

        if (verticalAxis.MaxAxisType == XLSparklineAxisMinMax.Custom)
        {
            WorksheetXml.SetOptional(sparklineGroup, "manualMax", verticalAxis.ManualMax);
        }

        IXLSparklineStyle style = xlSparklineGroup.Style;
        Color("colorSeries", style.SeriesColor);
        Color("colorNegative", style.NegativeColor);
        Color("colorAxis", horizontalAxis.Color);
        Color("colorMarkers", style.MarkersColor);
        Color("colorFirst", style.FirstMarkerColor);
        Color("colorLast", style.LastMarkerColor);
        Color("colorHigh", style.HighMarkerColor);
        Color("colorLow", style.LowMarkerColor);

        if (horizontalAxis.DateAxis)
        {
            sparklineGroup.Add(
                new XElement(
                    SpreadsheetXml.Xm + "f",
                    xlSparklineGroup.DateRange.RangeAddress.ToString(XLReferenceStyle.A1, true)
                )
            );
        }

        sparklineGroup.Add(
            new XElement(
                SpreadsheetXml.X14 + "sparklines",
                xlSparklineGroup.Sparklines.Select(xlSparkline => new XElement(
                    SpreadsheetXml.X14 + "sparkline",
                    // When sparkline source data area is deleted, Excel shows it as #REF! and is
                    // saved in file as an empty string
                    new XElement(
                        SpreadsheetXml.Xm + "f",
                        xlSparkline.SourceDataFormula ?? string.Empty
                    ),
                    new XElement(SpreadsheetXml.Xm + "sqref", xlSparkline.Location.ToString())
                ))
            )
        );

        return sparklineGroup;

        void Marker(string name, XLSparklineMarkers marker) =>
            WorksheetXml.SetBool(
                sparklineGroup,
                name,
                xlSparklineGroup.ShowMarkers.HasFlag(marker)
            );

        void Color(string name, XLColor color)
        {
            XElement element = new(SpreadsheetXml.X14 + name);
            SpreadsheetXml.SetColor(element, color);
            sparklineGroup.Add(element);
        }
    }

    /// <summary>
    /// The named extension's content element, made along with the extension and the list around
    /// it if the sheet has none, and emptied if it has.
    /// </summary>
    private static XElement Extension(XElement worksheet, string uri, string contentName)
    {
        XElement extensionList = WorksheetXml.Child(worksheet, "extLst");
        XElement content = extensionList
            .Descendants(SpreadsheetXml.X14 + contentName)
            .SingleOrDefault();
        if (content is not null && content.Elements().Any())
        {
            content.RemoveNodes();
            return content;
        }

        content = new XElement(
            SpreadsheetXml.X14 + contentName,
            new XAttribute(XNamespace.Xmlns + "xm", SpreadsheetXml.Xm.NamespaceName)
        );
        extensionList.Add(
            new XElement(
                SpreadsheetXml.Main + "ext",
                new XAttribute(XNamespace.Xmlns + "x14", SpreadsheetXml.X14.NamespaceName),
                new XAttribute("uri", uri),
                content
            )
        );
        return content;
    }

    /// <summary>
    /// Drops the named extension, and the list around it if nothing else is in it.
    /// </summary>
    private static void RemoveExtension(XElement worksheet, string uri, XName contentName)
    {
        XElement extensionList = worksheet.Element(SpreadsheetXml.Main + "extLst");
        XElement extension = extensionList
            ?.Elements(SpreadsheetXml.Main + "ext")
            .FirstOrDefault(candidate =>
                string.Equals(
                    SpreadsheetXml.String(candidate, "uri"),
                    uri,
                    StringComparison.OrdinalIgnoreCase
                )
            );

        extension?.Elements(contentName).Remove();
        if (extension is not null && !extension.Elements().Any())
        {
            extension.Remove();
        }

        if (extensionList is not null && !extensionList.Elements().Any())
        {
            extensionList.Remove();
        }
    }

    /// <summary>
    /// The sheet's data validations, which go to one of two places. A validation whose list or
    /// bounds point at another sheet cannot be said in the 2006 schema at all, so it is written
    /// in the x14 extension instead; the rest go in dataValidations where they belong.
    /// </summary>
    private static void WriteDataValidations(
        XElement worksheet,
        XLWorksheet xlWorksheet,
        SaveOptions options
    )
    {
        if (options.ConsolidateDataValidationRanges)
        {
            xlWorksheet.DataValidations.Consolidate();
        }

        List<(IXLDataValidation Validation, string MinValue, string MaxValue)> standard = [];
        List<(IXLDataValidation Validation, string MinValue, string MaxValue)> extension = [];
        foreach (XLDataValidation validation in xlWorksheet.DataValidations)
        {
            (bool minIsElsewhere, string minValue) = OnAnotherSheet(
                xlWorksheet,
                validation.MinValue
            );
            (bool maxIsElsewhere, string maxValue) = OnAnotherSheet(
                xlWorksheet,
                validation.MaxValue
            );

            (minIsElsewhere || maxIsElsewhere ? extension : standard).Add(
                (validation, minValue, maxValue)
            );
        }

        WriteStandardDataValidations(worksheet, standard);
        WriteExtensionDataValidations(worksheet, extension);

        // The spec wants a reference to a range on this sheet written without the sheet name,
        // and one to another sheet is what forces the validation into the extension.
        static (bool, string) OnAnotherSheet(XLWorksheet sheet, string value)
        {
            if (!XlsxSharp.XLHelper.IsValidRangeAddress(value))
            {
                return (false, value);
            }

            int separatorIndex = value.LastIndexOf('!');
            if (separatorIndex < 0)
            {
                return (false, value);
            }

            string sheetName = value[..separatorIndex].UnescapeSheetName();
            return XlsxSharp.XLHelper.SheetComparer.Equals(sheet.Name, sheetName)
                ? (false, value[(separatorIndex + 1)..])
                : (true, value);
        }
    }

    private static void WriteStandardDataValidations(
        XElement worksheet,
        List<(IXLDataValidation Validation, string MinValue, string MaxValue)> validations
    )
    {
        // The element must have at least one child, so a sheet whose validations all say nothing
        // has none at all.
        if (!validations.Any(validation => validation.Validation.IsDirty()))
        {
            worksheet.Element(SpreadsheetXml.Main + "dataValidations")?.Remove();
            return;
        }

        XElement dataValidations = WorksheetXml.Child(worksheet, "dataValidations");
        dataValidations.Elements(SpreadsheetXml.Main + "dataValidation").Remove();

        foreach ((IXLDataValidation validation, string minValue, string maxValue) in validations)
        {
            XElement element = new(SpreadsheetXml.Main + "dataValidation");
            SetDataValidationAttributes(element, validation);
            element.SetAttributeValue(
                "sqref",
                string.Join(" ", validation.Ranges.Select(range => range.RangeAddress))
            );
            element.Add(
                new XElement(SpreadsheetXml.Main + "formula1", minValue),
                new XElement(SpreadsheetXml.Main + "formula2", maxValue)
            );
            dataValidations.Add(element);
        }

        WorksheetXml.Set(dataValidations, "count", (uint)validations.Count);
    }

    private static void WriteExtensionDataValidations(
        XElement worksheet,
        List<(IXLDataValidation Validation, string MinValue, string MaxValue)> validations
    )
    {
        const string uri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
        XElement extensionList = worksheet.Element(SpreadsheetXml.Main + "extLst");

        if (validations.Count == 0)
        {
            // The extension the sheet was loaded with goes, and the list around it goes too if
            // nothing else is in it.
            XElement extension = extensionList
                ?.Elements(SpreadsheetXml.Main + "ext")
                .FirstOrDefault(candidate =>
                    string.Equals(
                        SpreadsheetXml.String(candidate, "uri"),
                        uri,
                        StringComparison.OrdinalIgnoreCase
                    )
                );

            extension?.Elements(SpreadsheetXml.X14 + "dataValidations").Remove();
            if (extension is not null && !extension.Elements().Any())
            {
                extension.Remove();
            }

            if (extensionList is not null && !extensionList.Elements().Any())
            {
                extensionList.Remove();
            }

            return;
        }

        extensionList = WorksheetXml.Child(worksheet, "extLst");
        XElement dataValidations = extensionList
            .Descendants(SpreadsheetXml.X14 + "dataValidations")
            .SingleOrDefault();
        if (dataValidations is null || !dataValidations.Elements().Any())
        {
            dataValidations = new XElement(
                SpreadsheetXml.X14 + "dataValidations",
                new XAttribute(XNamespace.Xmlns + "xm", SpreadsheetXml.Xm.NamespaceName)
            );
            extensionList.Add(
                new XElement(
                    SpreadsheetXml.Main + "ext",
                    new XAttribute(XNamespace.Xmlns + "x14", SpreadsheetXml.X14.NamespaceName),
                    new XAttribute("uri", uri),
                    dataValidations
                )
            );
        }
        else
        {
            dataValidations.RemoveNodes();
        }

        foreach ((IXLDataValidation validation, string minValue, string maxValue) in validations)
        {
            XElement element = new(SpreadsheetXml.X14 + "dataValidation");
            SetDataValidationAttributes(element, validation);
            Formula("formula1", minValue);
            Formula("formula2", maxValue);
            element.Add(
                new XElement(
                    SpreadsheetXml.Xm + "sqref",
                    string.Join(" ", validation.Ranges.Select(range => range.RangeAddress))
                )
            );
            dataValidations.Add(element);

            void Formula(string name, string value)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    element.Add(
                        new XElement(
                            SpreadsheetXml.X14 + name,
                            new XElement(SpreadsheetXml.Xm + "f", value)
                        )
                    );
                }
            }
        }

        WorksheetXml.Set(dataValidations, "count", (uint)validations.Count);
    }

    /// <summary>
    /// The attributes a data validation carries, which the 2006 schema and the x14 extension
    /// spell the same way.
    /// </summary>
    private static void SetDataValidationAttributes(XElement element, IXLDataValidation validation)
    {
        element.SetAttributeValue("type", validation.AllowedValues.ToXml());
        element.SetAttributeValue("errorStyle", validation.ErrorStyle.ToXml());
        element.SetAttributeValue("operator", validation.Operator.ToXml());
        WorksheetXml.SetBool(element, "allowBlank", validation.IgnoreBlanks);
        WorksheetXml.SetBool(element, "showDropDown", !validation.InCellDropdown);
        WorksheetXml.SetBool(element, "showInputMessage", validation.ShowInputMessage);
        WorksheetXml.SetBool(element, "showErrorMessage", validation.ShowErrorMessage);
        element.SetAttributeValue("errorTitle", validation.ErrorTitle);
        element.SetAttributeValue("error", validation.ErrorMessage);
        element.SetAttributeValue("promptTitle", validation.InputTitle);
        element.SetAttributeValue("prompt", validation.InputMessage);
    }

    /// <summary>
    /// The <c>conditionalFormatting</c> groups, one per set of ranges, and the x14 extension the
    /// data bars among them need.
    /// </summary>
    private static void WriteConditionalFormats(
        XElement worksheet,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        HashSet<XLConditionalFormat> pivotFormats =
        [
            .. xlWorksheet.PivotTables.SelectMany<XLPivotTable, XLConditionalFormat>(pivotTable =>
                pivotTable.ConditionalFormats.Select(cf => cf.Format)
            ),
        ];

        // Elements in sheet.ConditionalFormats were sorted according to priority during load,
        // but new ones have priority 0. CFs are also interleaved with sheet CF. To deal with
        // these situations, set correct unique priority (also required for pivot CF).
        List<XLConditionalFormat> formats =
        [
            .. xlWorksheet
                .ConditionalFormats.Cast<XLConditionalFormat>()
                .Concat(pivotFormats)
                .OrderBy(format => format.Priority),
        ];
        for (int i = 0; i < formats.Count; ++i)
        {
            formats[i].Priority = i + 1;
        }

        worksheet.Elements(SpreadsheetXml.Main + "conditionalFormatting").Remove();

        XElement previous = null;
        foreach (
            IGrouping<(string Ranges, bool IsPivot), XLConditionalFormat> group in formats.GroupBy(
                format => (Ranges(format), IsPivot: pivotFormats.Contains(format))
            )
        )
        {
            XElement conditionalFormatting = new(
                SpreadsheetXml.Main + "conditionalFormatting",
                new XAttribute("sqref", group.Key.Ranges),
                group.Select(format => ConditionalFormatXml.Rule(format, format.Priority, context))
            );

            if (group.Key.IsPivot)
            {
                WorksheetXml.SetBool(conditionalFormatting, "pivot", true);
            }

            if (previous is null)
            {
                WorksheetXml.Insert(worksheet, "conditionalFormatting", conditionalFormatting);
            }
            else
            {
                previous.AddAfterSelf(conditionalFormatting);
            }

            previous = conditionalFormatting;
        }

        WriteConditionalFormatExtensions(worksheet, xlWorksheet);

        static string Ranges(XLConditionalFormat format) =>
            string.Join(
                " ",
                format.Ranges.Select(range => range.RangeAddress.ToStringRelative(false))
            );
    }

    /// <summary>
    /// The x14 rules of the sheet's data bars, which carry the negative colour and the axis the
    /// 2006 schema has no place for. A rule already in the extension list is replaced along with
    /// the formatting around it, so a workbook that is loaded and saved keeps one of each.
    /// </summary>
    private static void WriteConditionalFormatExtensions(
        XElement worksheet,
        XLWorksheet xlWorksheet
    )
    {
        List<XLConditionalFormat> dataBars =
        [
            .. xlWorksheet.ConditionalFormats.Where<XLConditionalFormat>(format =>
                format.ConditionalFormatType == XLConditionalFormatType.DataBar
            ),
        ];
        if (dataBars.Count == 0)
        {
            return;
        }

        XElement extensionList = WorksheetXml.Child(worksheet, "extLst");
        XElement conditionalFormattings = extensionList
            .Descendants(SpreadsheetXml.X14 + "conditionalFormattings")
            .SingleOrDefault();
        if (conditionalFormattings is null || !conditionalFormattings.Elements().Any())
        {
            conditionalFormattings = new XElement(SpreadsheetXml.X14 + "conditionalFormattings");
            extensionList.Add(
                new XElement(
                    SpreadsheetXml.Main + "ext",
                    new XAttribute(XNamespace.Xmlns + "x14", SpreadsheetXml.X14.NamespaceName),
                    new XAttribute("uri", "{78C0D931-6437-407d-A8EE-F0AAD7539E65}"),
                    conditionalFormattings
                )
            );
        }

        foreach (XLConditionalFormat dataBar in dataBars)
        {
            string id = dataBar.Id.WrapInBraces();
            conditionalFormattings
                .Elements(SpreadsheetXml.X14 + "conditionalFormatting")
                .Where(formatting =>
                    formatting
                        .Elements(SpreadsheetXml.X14 + "cfRule")
                        .Any(rule => SpreadsheetXml.String(rule, "id") == id)
                )
                .Remove();

            conditionalFormattings.Add(
                new XElement(
                    SpreadsheetXml.X14 + "conditionalFormatting",
                    new XAttribute(XNamespace.Xmlns + "xm", SpreadsheetXml.Xm.NamespaceName),
                    ConditionalFormatXml.ExtensionRule(dataBar),
                    new XElement(
                        SpreadsheetXml.Xm + "sqref",
                        string.Join(
                            " ",
                            dataBar.Ranges.Select(range =>
                                range.RangeAddress.ToStringRelative(false)
                            )
                        )
                    )
                )
            );
        }
    }

    /// <summary>
    /// <c>sheetProtection</c>, which is written only for a protected sheet.
    /// </summary>
    private static void WriteSheetProtection(XElement worksheet, XLSheetProtection protection)
    {
        if (!protection.IsProtected)
        {
            worksheet.Element(SpreadsheetXml.Main + "sheetProtection")?.Remove();
            return;
        }

        XElement element = WorksheetXml.Child(worksheet, "sheetProtection");
        WorksheetXml.SetBoolDefault(element, "sheet", true, false);

        // The password is written one way or the other, never both, so the way not taken is
        // cleared off whatever the loaded sheet carried.
        element.SetAttributeValue("password", null);
        element.SetAttributeValue("algorithmName", null);
        element.SetAttributeValue("hashValue", null);
        element.SetAttributeValue("spinCount", null);
        element.SetAttributeValue("saltValue", null);

        if (protection.Algorithm == XLProtectionAlgorithm.Algorithm.SimpleHash)
        {
            if (!string.IsNullOrWhiteSpace(protection.PasswordHash))
            {
                element.SetAttributeValue("password", protection.PasswordHash);
            }
        }
        else
        {
            element.SetAttributeValue(
                "algorithmName",
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.ToDescription(
                    protection.Algorithm
                )
            );
            element.SetAttributeValue("hashValue", protection.PasswordHash);
            WorksheetXml.Set(element, "spinCount", protection.SpinCount);
            element.SetAttributeValue("saltValue", protection.Base64EncodedSalt);
        }

        // Every attribute says what is denied, so an element the sheet allows turns its attribute
        // off. They differ only in what the schema already says about them.
        Deny(XLSheetProtectionElements.FormatCells, "formatCells", true);
        Deny(XLSheetProtectionElements.FormatColumns, "formatColumns", true);
        Deny(XLSheetProtectionElements.FormatRows, "formatRows", true);
        Deny(XLSheetProtectionElements.InsertColumns, "insertColumns", true);
        Deny(XLSheetProtectionElements.InsertRows, "insertRows", true);
        Deny(XLSheetProtectionElements.InsertHyperlinks, "insertHyperlinks", true);
        Deny(XLSheetProtectionElements.DeleteColumns, "deleteColumns", true);
        Deny(XLSheetProtectionElements.DeleteRows, "deleteRows", true);
        Deny(XLSheetProtectionElements.Sort, "sort", true);
        Deny(XLSheetProtectionElements.AutoFilter, "autoFilter", true);
        Deny(XLSheetProtectionElements.PivotTables, "pivotTables", true);
        Deny(XLSheetProtectionElements.EditScenarios, "scenarios", true);
        Deny(XLSheetProtectionElements.EditObjects, "objects", false);
        Deny(XLSheetProtectionElements.SelectLockedCells, "selectLockedCells", false);
        Deny(XLSheetProtectionElements.SelectUnlockedCells, "selectUnlockedCells", false);

        void Deny(XLSheetProtectionElements allowed, string name, bool deniedByDefault) =>
            WorksheetXml.SetBoolDefault(
                element,
                name,
                !protection.AllowedElements.HasFlag(allowed),
                deniedByDefault
            );
    }

    /// <summary>
    /// <c>mergeCells</c>, which is written whole from the workbook model.
    /// </summary>
    private static void WriteMergedCells(XElement worksheet, XLRanges mergedRanges)
    {
        if (mergedRanges.Count == 0)
        {
            worksheet.Element(SpreadsheetXml.Main + "mergeCells")?.Remove();
            return;
        }

        XElement element = WorksheetXml.Child(worksheet, "mergeCells");
        element.RemoveNodes();
        foreach (XLRange range in mergedRanges)
        {
            element.Add(
                new XElement(
                    SpreadsheetXml.Main + "mergeCell",
                    new XAttribute(
                        "ref",
                        $"{range.RangeAddress.FirstAddress}:{range.RangeAddress.LastAddress}"
                    )
                )
            );
        }

        WorksheetXml.Set(element, "count", (uint)mergedRanges.Count);
    }

    /// <summary>
    /// <c>hyperlinks</c>, with a relationship for each link that points outside the workbook.
    /// The relationships of the sheet's previous links go first, so a link removed from the
    /// workbook takes its relationship with it.
    /// </summary>
    private static void WriteHyperlinks(
        XElement worksheet,
        OpcPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        foreach (
            string relationshipId in worksheetPart
                .Relationships.OfType(OoxmlPartTypes.HyperlinkRelationshipType)
                .Select(r => r.Id)
                .ToList()
        )
        {
            worksheetPart.Relationships.Remove(relationshipId);
        }

        if (!xlWorksheet.Hyperlinks.Any())
        {
            worksheet.Element(SpreadsheetXml.Main + "hyperlinks")?.Remove();
            return;
        }

        XElement element = WorksheetXml.Child(worksheet, "hyperlinks");
        element.RemoveNodes();
        foreach (XLHyperlink hyperlink in xlWorksheet.Hyperlinks)
        {
            XElement written = new(
                SpreadsheetXml.Main + "hyperlink",
                new XAttribute("ref", hyperlink.Cell.Address.ToString())
            );

            if (hyperlink.IsExternal)
            {
                string relId = context.RelIdGenerator.GetNext(XLWorkbook.RelType.Workbook);
                written.SetAttributeValue(SpreadsheetXml.Rel + "id", relId);
                worksheetPart.Relationships.AddExternal(
                    hyperlink.ExternalAddress.OriginalString,
                    OoxmlPartTypes.HyperlinkRelationshipType,
                    relId
                );
            }
            else
            {
                written.SetAttributeValue("location", hyperlink.InternalAddress);
                written.SetAttributeValue("display", hyperlink.Cell.GetFormattedString());
            }

            if (!string.IsNullOrWhiteSpace(hyperlink.Tooltip))
            {
                written.SetAttributeValue("tooltip", hyperlink.Tooltip);
            }

            element.Add(written);
        }
    }

    /// <summary>
    /// <c>printOptions</c>, <c>pageMargins</c> and <c>pageSetup</c>, which say how the sheet is
    /// printed. All three are always written, and all three replace whatever the loaded sheet
    /// carried - the workbook model holds every one of their attributes.
    /// </summary>
    private static void WritePageSetup(XElement worksheet, IXLPageSetup pageSetup)
    {
        XElement printOptions = WorksheetXml.Child(worksheet, "printOptions");
        WorksheetXml.SetBool(printOptions, "horizontalCentered", pageSetup.CenterHorizontally);
        WorksheetXml.SetBool(printOptions, "verticalCentered", pageSetup.CenterVertically);
        WorksheetXml.SetBool(printOptions, "headings", pageSetup.ShowRowAndColumnHeadings);
        WorksheetXml.SetBool(printOptions, "gridLines", pageSetup.ShowGridlines);

        XElement margins = WorksheetXml.Child(worksheet, "pageMargins");
        WorksheetXml.Set(margins, "left", pageSetup.Margins.Left);
        WorksheetXml.Set(margins, "right", pageSetup.Margins.Right);
        WorksheetXml.Set(margins, "top", pageSetup.Margins.Top);
        WorksheetXml.Set(margins, "bottom", pageSetup.Margins.Bottom);
        WorksheetXml.Set(margins, "header", pageSetup.Margins.Header);
        WorksheetXml.Set(margins, "footer", pageSetup.Margins.Footer);

        WritePageSetupElement(WorksheetXml.Child(worksheet, "pageSetup"), pageSetup);
    }

    private static void WritePageSetupElement(XElement element, IXLPageSetup pageSetup)
    {
        WorksheetXml.Set(element, "paperSize", (uint)pageSetup.PaperSize);
        element.SetAttributeValue("orientation", pageSetup.PageOrientation.ToXml());
        element.SetAttributeValue("pageOrder", pageSetup.PageOrder.ToXml());
        element.SetAttributeValue("cellComments", pageSetup.ShowComments.ToXml());
        element.SetAttributeValue("errors", pageSetup.PrintErrorValue.ToXml());
        WorksheetXml.SetBool(element, "blackAndWhite", pageSetup.BlackAndWhite);
        WorksheetXml.SetBool(element, "draft", pageSetup.DraftQuality);

        if (pageSetup.FirstPageNumber is { } firstPageNumber)
        {
            // Negative first page numbers are written as uint, e.g. -1 is 4294967295.
            WorksheetXml.Set(element, "firstPageNumber", (uint)firstPageNumber);
            WorksheetXml.SetBool(element, "useFirstPageNumber", true);
        }
        else
        {
            element.SetAttributeValue("firstPageNumber", null);
            element.SetAttributeValue("useFirstPageNumber", null);
        }

        WorksheetXml.SetOptional<uint>(
            element,
            "horizontalDpi",
            pageSetup.HorizontalDpi > 0 ? (uint)pageSetup.HorizontalDpi : null
        );
        WorksheetXml.SetOptional<uint>(
            element,
            "verticalDpi",
            pageSetup.VerticalDpi > 0 ? (uint)pageSetup.VerticalDpi : null
        );

        // A sheet is either scaled or fitted to a number of pages, never both, and a count of one
        // page is the default the attribute exists to override.
        if (pageSetup.Scale > 0)
        {
            WorksheetXml.Set(element, "scale", (uint)pageSetup.Scale);
            element.SetAttributeValue("fitToWidth", null);
            element.SetAttributeValue("fitToHeight", null);
        }
        else
        {
            element.SetAttributeValue("scale", null);
            WorksheetXml.SetOptional<uint>(
                element,
                "fitToWidth",
                pageSetup.PagesWide >= 0 && pageSetup.PagesWide != 1
                    ? (uint)pageSetup.PagesWide
                    : null
            );
            WorksheetXml.SetOptional<uint>(
                element,
                "fitToHeight",
                pageSetup.PagesTall >= 0 && pageSetup.PagesTall != 1
                    ? (uint)pageSetup.PagesTall
                    : null
            );
        }

        // For some reason some Excel files already contains copies="0", which the validator
        // refuses. Drop the attribute when that is the case.
        if (SpreadsheetXml.UInt(element, "copies") is null or 0)
        {
            element.SetAttributeValue("copies", null);
        }
    }

    /// <summary>
    /// <c>headerFooter</c>, which is always written even when it says nothing - an empty element
    /// is what a sheet with untouched headers gets. Only a header or footer the workbook model
    /// has changed replaces what the loaded sheet carried.
    /// </summary>
    private static void WriteHeaderFooter(XElement worksheet, IXLPageSetup pageSetup)
    {
        XElement headerFooter = WorksheetXml.Child(worksheet, "headerFooter");
        if (
            !((XLHeaderFooter)pageSetup.Header).Changed
            && !((XLHeaderFooter)pageSetup.Footer).Changed
        )
        {
            return;
        }

        headerFooter.RemoveNodes();

        WorksheetXml.SetBool(headerFooter, "scaleWithDoc", pageSetup.ScaleHFWithDocument);
        WorksheetXml.SetBool(headerFooter, "alignWithMargins", pageSetup.AlignHFWithMargins);
        WorksheetXml.SetBool(headerFooter, "differentFirst", pageSetup.DifferentFirstPageOnHF);
        WorksheetXml.SetBool(headerFooter, "differentOddEven", pageSetup.DifferentOddEvenPagesOnHF);

        Text("oddHeader", pageSetup.Header, XLHFOccurrence.OddPages);
        Text("oddFooter", pageSetup.Footer, XLHFOccurrence.OddPages);
        Text("evenHeader", pageSetup.Header, XLHFOccurrence.EvenPages);
        Text("evenFooter", pageSetup.Footer, XLHFOccurrence.EvenPages);
        Text("firstHeader", pageSetup.Header, XLHFOccurrence.FirstPage);
        Text("firstFooter", pageSetup.Footer, XLHFOccurrence.FirstPage);

        void Text(string name, IXLHeaderFooter source, XLHFOccurrence occurrence) =>
            headerFooter.Add(new XElement(SpreadsheetXml.Main + name, source.GetText(occurrence)));
    }

    /// <summary>
    /// A list of manual page breaks. The breaks the sheet already carries are left where they
    /// are, so a file that names a break's width or its first row keeps saying so; only the ones
    /// the workbook model no longer has go, and the ones it has gained are added at the end.
    /// </summary>
    private static void WriteBreaks(
        XElement worksheet,
        string name,
        List<int> breaks,
        uint lastLine
    )
    {
        if (breaks.Count == 0)
        {
            worksheet.Element(SpreadsheetXml.Main + name)?.Remove();
            return;
        }

        XElement element = WorksheetXml.Child(worksheet, name);
        List<uint> kept = [];
        foreach (XElement brk in element.Elements(SpreadsheetXml.Main + "brk").ToList())
        {
            if (SpreadsheetXml.UInt(brk, "id") is { } id && breaks.Contains(checked((int)id)))
            {
                kept.Add(id);
            }
            else
            {
                brk.Remove();
            }
        }

        WorksheetXml.Set(element, "count", (uint)breaks.Count);
        WorksheetXml.Set(element, "manualBreakCount", (uint)breaks.Count);

        foreach (int id in breaks.Where(id => !kept.Contains((uint)id)))
        {
            element.Add(
                new XElement(
                    SpreadsheetXml.Main + "brk",
                    new XAttribute("id", id),
                    new XAttribute("max", lastLine),
                    new XAttribute("man", "1")
                )
            );
        }
    }

    private static void StreamToPart(
        XElement worksheet,
        OpcPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context,
        SaveOptions options
    )
    {
        // Worksheet part might have some data, but creating the stream truncates everything.
        using Stream partStream = worksheetPart.GetWriteStream();
        using XmlWriter xml = XmlWriter.Create(
            partStream,
            new XmlWriterSettings { CloseOutput = true, Encoding = XlsxSharp.XLHelper.NoBomUTF8 }
        );

        xml.WriteStartDocument(true);
        xml.WriteStartElement(
            worksheet.GetPrefixOfNamespace(worksheet.Name.Namespace),
            worksheet.Name.LocalName,
            worksheet.Name.NamespaceName
        );

        foreach (XAttribute attribute in worksheet.Attributes())
        {
            WriteAttribute(xml, worksheet, attribute);
        }

        foreach (XElement child in worksheet.Elements())
        {
            if (child.Name == SpreadsheetXml.Main + "sheetData")
            {
                StreamSheetData(xml, xlWorksheet, context, options);
            }
            else
            {
                child.WriteTo(xml);
            }
        }

        xml.WriteEndElement();
        xml.WriteEndDocument();
    }

    private static void WriteAttribute(XmlWriter xml, XElement element, XAttribute attribute)
    {
        if (attribute.IsNamespaceDeclaration)
        {
            // A prefixed declaration is an attribute in the xmlns namespace; the default one is
            // an attribute called xmlns in no namespace at all.
            if (attribute.Name.Namespace == XNamespace.Xmlns)
            {
                xml.WriteAttributeString(
                    "xmlns",
                    attribute.Name.LocalName,
                    XNamespace.Xmlns.NamespaceName,
                    attribute.Value
                );
            }
            else
            {
                xml.WriteAttributeString("xmlns", attribute.Value);
            }

            return;
        }

        xml.WriteAttributeString(
            attribute.Name.Namespace == XNamespace.None
                ? null
                : element.GetPrefixOfNamespace(attribute.Name.Namespace),
            attribute.Name.LocalName,
            attribute.Name.NamespaceName,
            attribute.Value
        );
    }

    private static void StreamSheetData(
        XmlWriter xml,
        XLWorksheet xlWorksheet,
        SaveContext context,
        SaveOptions options
    )
    {
        int maxColumn = GetMaxColumn(xlWorksheet);

        xml.WriteStartElement("sheetData", Main2006SsNs);

        HashSet<IXLAddress> tableTotalCells =
        [
            .. xlWorksheet
                .Tables.Where<XLTable>(table => table.ShowTotalsRow)
                .SelectMany(table => table.TotalsRow().CellsUsed())
                .Select(cell => cell.Address),
        ];

        // A rather complicated state machine, so rows and cells can be written in a single loop
        int openedRowNumber = 0;
        bool isRowOpened = false;
        char[] cellRef = new char[10]; // Buffer, must be enough to hold span and rowNumber as strings
        List<int> rows = [.. xlWorksheet.Internals.RowsCollection.Keys];
        rows.Sort();
        int rowPropIndex = 0;
        uint rowStyleId = 0;
        foreach (XLCell xlCell in xlWorksheet.Internals.CellsCollection.GetCells())
        {
            int currentRowNumber = xlCell.Point.Row;

            // A space between cells can have several rows that don't contain cells,
            // but have custom properties (e.g. height). Write them out.
            while (rowPropIndex < rows.Count && rows[rowPropIndex] < currentRowNumber)
            {
                if (isRowOpened)
                {
                    xml.WriteEndElement(); // row
                    isRowOpened = false;
                }

                int rowNumber = rows[rowPropIndex];
                XLRow xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
                if (RowHasCustomProps(xlRow))
                {
                    WriteStartRow(xml, xlRow, rowNumber, maxColumn, context);

                    isRowOpened = true;
                    openedRowNumber = rowNumber;
                }

                rowPropIndex++;
            }

            // For saving cells to file, ignore conditional formatting, data validation rules and merged
            // ranges. They just bloat the file
            bool isEmpty =
                xlCell.CachedValue.Type == XLDataType.Blank
                && xlCell.IsEmpty(
                    XLCellsUsedOptions.All
                        & ~XLCellsUsedOptions.ConditionalFormats
                        & ~XLCellsUsedOptions.DataValidation
                        & ~XLCellsUsedOptions.MergedRanges
                );

            if (isEmpty)
            {
                continue;
            }

            if (openedRowNumber != currentRowNumber)
            {
                if (isRowOpened)
                {
                    xml.WriteEndElement(); // row
                }

                if (
                    xlWorksheet.Internals.RowsCollection.TryGetValue(
                        currentRowNumber,
                        out XLRow row
                    )
                )
                {
                    rowPropIndex++;
                    rowStyleId = context.GetStyleId(row.FormatValue);
                }
                else
                {
                    rowStyleId = 0;
                }

                WriteStartRow(xml, row, currentRowNumber, maxColumn, context);

                isRowOpened = true;
                openedRowNumber = currentRowNumber;
            }

            WriteCell(xml, xlCell, cellRef, context, options, tableTotalCells, rowStyleId);
        }

        if (isRowOpened)
        {
            xml.WriteEndElement(); // row
        }

        // Write rows with custom properties after last cell.
        while (rowPropIndex < rows.Count)
        {
            int rowNumber = rows[rowPropIndex];
            XLRow xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
            if (RowHasCustomProps(xlRow))
            {
                WriteStartRow(xml, xlRow, rowNumber, 0, context);
                xml.WriteEndElement(); // row
            }

            rowPropIndex++;
        }

        xml.WriteEndElement(); // SheetData

        static bool RowHasCustomProps(XLRow xlRow)
        {
            return xlRow.HeightChanged
                || xlRow.IsHidden
                || xlRow.FormatValue is not null
                || xlRow.Collapsed
                || xlRow.OutlineLevel > 0;
        }

        static void WriteStartRow(
            XmlWriter w,
            XLRow xlRow,
            int rowNumber,
            int maxColumn,
            SaveContext context
        )
        {
            w.WriteStartElement("row", Main2006SsNs);

            w.WriteStartAttribute("r");
            w.WriteValue(rowNumber);
            w.WriteEndAttribute();

            if (maxColumn > 0)
            {
                w.WriteStartAttribute("spans");
                w.WriteString("1:");
                w.WriteValue(maxColumn);
                w.WriteEndAttribute();
            }

            if (xlRow is null)
            {
                return;
            }

            if (xlRow.HeightChanged)
            {
                double height = xlRow.Height.SaveRound();
                w.WriteStartAttribute("ht");
                w.WriteNumberValue(height);
                w.WriteEndAttribute();

                // Note that dyDescent automatically implies custom height
                w.WriteAttributeString("customHeight", TrueValue);
            }

            if (xlRow.IsHidden)
            {
                w.WriteAttributeString("hidden", TrueValue);
            }

            bool rowHasCustomFormat = xlRow.FormatValue is not null;
            if (rowHasCustomFormat)
            {
                uint formatIndex = context.GetStyleId(xlRow.FormatValue);
                w.WriteAttribute("s", formatIndex);
                w.WriteAttributeString("customFormat", TrueValue);
            }

            if (xlRow.Collapsed)
            {
                w.WriteAttributeString("collapsed", TrueValue);
            }

            if (xlRow.OutlineLevel > 0)
            {
                w.WriteAttribute("outlineLevel", xlRow.OutlineLevel);
            }

            if (xlRow.ShowPhonetic)
            {
                w.WriteAttributeString("ph", TrueValue);
            }

            if (xlRow.DyDescent is not null)
            {
                w.WriteAttribute("dyDescent", X14Ac2009SsNs, xlRow.DyDescent.Value);
            }

            // thickBot and thickTop attributes are not written, because Excel seems to determine adjustments
            // from cell borders on its own and it would be rather costly to check each cell in each row.
            // If row was adjusted when cell had it's border modified, then it would be fine to write
            // the thickBot/thickBot attributes.
        }

        static void WriteStartCell(
            XmlWriter w,
            XLCell xlCell,
            char[] reference,
            int referenceLength,
            string dataType,
            uint styleId
        )
        {
            w.WriteStartElement("c", Main2006SsNs);

            w.WriteStartAttribute("r");
            w.WriteRaw(reference, 0, referenceLength);
            w.WriteEndAttribute();

            // TODO: if (styleId != 0) Test files have style even for 0, fix later
            w.WriteAttribute("s", styleId);

            if (dataType is not null)
            {
                w.WriteAttributeString("t", dataType);
            }

            if (xlCell.ShowPhonetic)
            {
                w.WriteAttributeString("ph", TrueValue);
            }

            if (xlCell.CellMetaIndex is not null)
            {
                w.WriteAttribute("cm", xlCell.CellMetaIndex.Value);
            }

            if (xlCell.ValueMetaIndex is not null)
            {
                w.WriteAttribute("vm", xlCell.ValueMetaIndex.Value);
            }
        }

        static void WriteCell(
            XmlWriter xml,
            XLCell xlCell,
            char[] cellRef,
            SaveContext context,
            SaveOptions options,
            HashSet<IXLAddress> tableTotalCells,
            uint rowStyleId
        )
        {
            uint styleId = context.GetStyleId(xlCell.GetFormat());

            Span<char> cellRefSpan = cellRef;
            int cellRefLen = xlCell.Point.Format(cellRefSpan);

            if (xlCell.HasFormula)
            {
                string dataType = null;
                if (options.EvaluateFormulasBeforeSaving)
                {
                    try
                    {
                        xlCell.Evaluate(false);
                        dataType = FormulaDataType[(int)xlCell.DataType];
                    }
                    catch
                    {
                        // Do nothing, cell will be left blank. Unimplemented features or functions would stop trying to save a file.
                    }
                }

                WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                XLCellFormula xlFormula = xlCell.Formula;
                if (xlFormula.Type == FormulaType.DataTable)
                {
                    // Data table doesn't write actual text of formula, that is referenced by context
                    xml.WriteStartElement("f", Main2006SsNs);
                    xml.WriteAttributeString("t", "dataTable");
                    xml.WriteAttributeString("ref", xlFormula.Range.ToString());

                    bool is2D = xlFormula.Is2DDataTable;
                    if (is2D)
                    {
                        xml.WriteAttributeString("dt2D", TrueValue);
                    }

                    bool isDataRowTable = xlFormula.IsRowDataTable;
                    if (isDataRowTable)
                    {
                        xml.WriteAttributeString("dtr", TrueValue);
                    }

                    xml.WriteAttributeString("r1", xlFormula.Input1.ToString());
                    bool input1Deleted = xlFormula.Input1Deleted;
                    if (input1Deleted)
                    {
                        xml.WriteAttributeString("del1", TrueValue);
                    }

                    if (is2D)
                    {
                        xml.WriteAttributeString("r2", xlFormula.Input2.ToString());
                    }

                    bool input2Deleted = xlFormula.Input2Deleted;
                    if (input2Deleted)
                    {
                        xml.WriteAttributeString("del2", TrueValue);
                    }

                    // Excel doesn't recalculate table formula on load or on click of a button or any kind of forced recalculation.
                    // It is necessary to mark some precedent formula dirty (e.g. edit cell formula and enter in Excel).
                    // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
                    // user will see correct values.
                    xml.WriteAttributeString("ca", TrueValue);

                    xml.WriteEndElement(); // f
                }
                else if (xlCell.HasArrayFormula)
                {
                    bool isMasterCell = xlCell.Formula.Range.FirstPoint == xlCell.Point;
                    if (isMasterCell)
                    {
                        xml.WriteStartElement("f", Main2006SsNs);
                        xml.WriteAttributeString("t", "array");
                        xml.WriteAttributeString("ref", xlCell.FormulaReference.ToStringRelative());
                        xml.WriteString(xlCell.FormulaA1);
                        xml.WriteEndElement(); // f
                    }
                }
                else
                {
                    xml.WriteStartElement("f", Main2006SsNs);
                    xml.WriteString(xlCell.FormulaA1);
                    xml.WriteEndElement(); // f
                }

                if (
                    options.EvaluateFormulasBeforeSaving
                    && xlCell.CachedValue.Type != XLDataType.Blank
                    && !xlCell.NeedsRecalculation
                )
                {
                    WriteCellValue(xml, xlCell, context);
                }

                xml.WriteEndElement(); // cell
            }
            else if (tableTotalCells.Contains(xlCell.Address))
            {
                XLTable table = xlCell.Worksheet.Tables.First<XLTable>(t =>
                    t.AsRange().Contains(xlCell)
                );
                XLTableField field = (XLTableField)
                    table.Fields.First(f => f.Column.ColumnNumber() == xlCell.Address.ColumnNumber);

                // If this is a cell in the totals row that contains a label (xor with function), write label
                // Only label can be written. Total functions are basically formulas that use structured
                // references and SR are not yet supported, so not yet possible to calculate total values.
                if (!string.IsNullOrWhiteSpace(field.TotalsRowLabel))
                {
                    // Excel requires that table totals row label attribute in tableColumn must match the cell
                    // string from SST. If they don't match, Excel will consider it a corrupt workbook.
                    int sharedStringId = context.GetSharedStringId(xlCell, field.TotalsRowLabel);
                    WriteStartCell(xml, xlCell, cellRef, cellRefLen, "s", styleId);
                    xml.WriteStartElement("v", Main2006SsNs);
                    xml.WriteValue(sharedStringId);
                    xml.WriteEndElement();
                }
                xml.WriteEndElement(); // cell
            }
            else if (xlCell.DataType != XLDataType.Blank)
            {
                // Cell contains only a value
                string dataType = GetCellValueType(xlCell);
                WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                WriteCellValue(xml, xlCell, context);
                xml.WriteEndElement(); // cell
            }
            else if (rowStyleId != styleId)
            {
                // Cell is blank and should be written only if it has different style from parent.
                // Non-written cells use inherited style of a row.
                WriteStartCell(xml, xlCell, cellRef, cellRefLen, null, styleId);
                xml.WriteEndElement(); // cell
            }
        }
    }

    /// <summary>
    /// An array to convert data type for a formula cell. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string[] FormulaDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "str", // text, formula can only save this type, no inline or shared string
        "e", // error
        null, // datetime, saved as serialized date-time
        null, // timespan, saved as serialized date-time
    ];

    /// <summary>
    /// An array to convert data type for a cell that only contains a value. Key is <see cref="XLDataType"/>.
    /// It saves some performance through direct indexation instead of switch.
    /// </summary>
    private static readonly string[] ValueDataType =
    [
        null, // blank
        "b", // boolean
        null, // number, default value, no need to save type
        "s", // text, the default is a shared string, but there also can be inline string depending on ShareString property
        "e", // error
        null, // datetime, saved as serialized date-time
        null, // timespan, saved as serialized date-time
    ];

    private static string GetCellValueType(XLCell xlCell)
    {
        XLDataType dataType = xlCell.DataType;
        if (dataType == XLDataType.Text && !xlCell.ShareString)
        {
            return "inlineStr";
        }

        return ValueDataType[(int)dataType];
    }

    private static int GetMaxColumn(XLWorksheet xlWorksheet)
    {
        int maxColumn = 0;

        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
        }

        if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
        {
            int maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
            if (maxColCollection > maxColumn)
            {
                maxColumn = maxColCollection;
            }
        }

        return maxColumn;
    }
}

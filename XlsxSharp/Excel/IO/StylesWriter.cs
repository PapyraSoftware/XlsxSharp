using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.Utils;
using static XlsxSharp.Excel.IO.OpenXmlConst;
using PivotRegion = XlsxSharp.Excel.Formatting.XLPivotStyleRegionValues;
using TableRegion = XlsxSharp.Excel.Formatting.XLTableStyleRegionValues;

namespace XlsxSharp.Excel.IO;

internal class StylesWriter
{
    private const int FirstUserDefinedFormatIndex =
        XLWorkbookStyles.FirstUserDefinedNumberFormatIndex;

    private static readonly List<(
        string Type,
        TableRegion? TableRegion,
        PivotRegion? PivotRegion
    )> TableRegionsMap =
    [
        ("wholeTable", TableRegion.WholeTable, PivotRegion.WholeTable),
        ("headerRow", TableRegion.HeaderRow, PivotRegion.HeaderRow),
        ("totalRow", TableRegion.TotalRow, PivotRegion.GrandTotalRow),
        ("firstColumn", TableRegion.FirstColumn, PivotRegion.FirstColumn),
        ("lastColumn", TableRegion.LastColumn, PivotRegion.GrandTotalColumn),
        ("firstRowStripe", TableRegion.FirstRowStripe, PivotRegion.FirstRowStripe),
        ("secondRowStripe", TableRegion.SecondRowStripe, PivotRegion.SecondRowStripe),
        ("firstColumnStripe", TableRegion.FirstColumnStripe, PivotRegion.FirstColumnStripe),
        ("secondColumnStripe", TableRegion.SecondColumnStripe, PivotRegion.SecondColumnStripe),
        ("firstHeaderCell", TableRegion.FirstHeaderCell, PivotRegion.FirstHeaderCell),
        ("lastHeaderCell", TableRegion.LastHeaderCell, null),
        ("firstTotalCell", TableRegion.FirstTotalCell, null),
        ("lastTotalCell", TableRegion.LastTotalCell, null),
        ("firstSubtotalColumn", null, PivotRegion.SubtotalColumn1),
        ("secondSubtotalColumn", null, PivotRegion.SubtotalColumn2),
        ("thirdSubtotalColumn", null, PivotRegion.SubtotalColumn3),
        ("firstSubtotalRow", null, PivotRegion.SubtotalRow1),
        ("secondSubtotalRow", null, PivotRegion.SubtotalRow2),
        ("thirdSubtotalRow", null, PivotRegion.SubtotalRow3),
        ("blankRow", null, PivotRegion.BlankRow),
        ("firstColumnSubheading", null, PivotRegion.ColumnSubheading1),
        ("secondColumnSubheading", null, PivotRegion.ColumnSubheading2),
        ("thirdColumnSubheading", null, PivotRegion.ColumnSubheading3),
        ("firstRowSubheading", null, PivotRegion.RowSubheading1),
        ("secondRowSubheading", null, PivotRegion.RowSubheading2),
        ("thirdRowSubheading", null, PivotRegion.RowSubheading3),
        ("pageFieldLabels", null, PivotRegion.PageFieldLabels),
        ("pageFieldValues", null, PivotRegion.PageFieldValues),
    ];

    private readonly string _ns = Main2006SsNs;

    internal void WriteContent(
        WorkbookStylesPart stylesPart,
        IEnumMapper mapper,
        XLWorkbookStyles styles,
        XLWorkbook workbook,
        XLWorkbook.SaveContext context
    )
    {
        // Determine which format components are used and thus should be saved.
        HashSet<XLCellFormatValue> usedCellFormats = new(
            ReferenceEqualityComparer<XLCellFormatValue>.Instance
        )
        {
            styles.DefaultCellFormat,
        };
        foreach (XLWorksheet sheet in workbook.WorksheetsInternal)
        {
            if (sheet.FormatValue is { } sheetFormat)
            {
                usedCellFormats.Add(sheetFormat);
            }

            foreach (XLColumn column in sheet.Internals.ColumnsCollection.Values)
            {
                if (column.FormatValue is { } columnFormat)
                {
                    usedCellFormats.Add(columnFormat);
                }
            }

            foreach (XLRow row in sheet.Internals.RowsCollection.Values)
            {
                if (row.FormatValue is { } rowFormat)
                {
                    usedCellFormats.Add(rowFormat);
                }
            }

            sheet.Internals.CellsCollection.FormatSlice.AddUsedFormat(usedCellFormats);
        }

        HashSet<XLNumberFormat> usedNumberFormats = [];
        HashSet<XLFontFormatValue> usedFonts = [];
        HashSet<XLFillFormatValue> usedFills = [];
        HashSet<XLBorderFormatValue> usedBorders = [];
        foreach (XLCellFormatValue cellFormat in usedCellFormats)
        {
            if (cellFormat.NumberFormat is { } numberFormat)
            {
                usedNumberFormats.Add(numberFormat);
            }

            if (cellFormat.Font is { } font)
            {
                usedFonts.Add(font);
            }

            if (cellFormat.Fill is { } fill)
            {
                usedFills.Add(fill);
            }

            if (cellFormat.Border is { } border)
            {
                usedBorders.Add(border);
            }
        }

        foreach (XLCellStyleValue cellStyle in styles.CellStyles.Values)
        {
            if (cellStyle.NumberFormat is { } numberFormat)
            {
                usedNumberFormats.Add(numberFormat);
            }

            if (cellStyle.Font is { } font)
            {
                usedFonts.Add(font);
            }

            if (cellStyle.Fill is { } fill)
            {
                usedFills.Add(fill);
            }

            if (cellStyle.Border is { } border)
            {
                usedBorders.Add(border);
            }
        }

        // SST writes fonts as ids, unlike other format properties which are inlined next to a rich text
        foreach (
            XLFontFormatValue phoneticsFont in workbook.SharedStringTable.GetUsedPhoneticFonts()
        )
        {
            usedFonts.Add(phoneticsFont);
        }

        // Create dxfMap from used dxfs in cfs, tables, pivot tables and so on
        HashSet<XLDxfValue> usedDxf = [];
        foreach (XLWorksheet ws in workbook.WorksheetsInternal)
        {
            foreach (XLConditionalFormat cf in ws.ConditionalFormats)
            {
                if (cf.FormatValue is { } dxf)
                {
                    usedDxf.Add(dxf);
                }
            }

            // Table styles should be retained in a workbook, even if not used in a table.
            foreach (XLTableTheme tableStyle in styles.TableStyles.Values)
            {
                foreach (XLDxfValue regionDxf in tableStyle.RegionFormats.Values)
                {
                    usedDxf.Add(regionDxf);
                }
            }

            foreach (XLTable table in ws.Tables)
            {
                foreach (XLTableField field in table.Fields)
                {
                    if (field.HeaderFormatValue is { } headerDxf)
                    {
                        usedDxf.Add(headerDxf);
                    }

                    if (field.DataFormatValue is { } dataDxf)
                    {
                        usedDxf.Add(dataDxf);
                    }

                    if (field.TotalFormatValue is { } totalsDxf)
                    {
                        usedDxf.Add(totalsDxf);
                    }
                }
            }

            foreach (XLPivotTable pt in ws.PivotTables)
            {
                foreach (XLPivotFormat format in pt.Formats)
                {
                    if (format.FormatValue is { } dxFormat)
                    {
                        usedDxf.Add(dxFormat);
                    }
                }

                foreach (XLPivotConditionalFormat cf in pt.ConditionalFormats)
                {
                    if (cf.Format.FormatValue is { } cfDxf)
                    {
                        usedDxf.Add(cfDxf);
                    }
                }

                foreach (XLPivotTableField field in pt.PivotFields)
                {
                    if (field.NumberFormatValue is { } fieldNumFmt)
                    {
                        usedNumberFormats.Add(fieldNumFmt);
                    }
                }
            }
        }

        XmlWriterSettings settings = new() { Encoding = XlsxSharp.XLHelper.NoBomUTF8 };

        using Stream partStream = stylesPart.GetStream(FileMode.Create);
        using XmlTreeWriter xml = new(XmlWriter.Create(partStream, settings), mapper);

        xml.WriteStartDocument("styleSheet", this._ns);

        // Number formats
        // The map has predefined formats from index 0 and the user defined ones from 164 onward.
        // There is a gap between predefined formats.
        IReadOnlyDictionary<XLNumberFormat, int> predefinedNumberFormats =
            XLPredefinedFormat.NumberFormatIds;
        SequentialMap<int, XLNumberFormat> numberFormatMap = SequentialMap<
            int,
            XLNumberFormat
        >.Create(
            usedNumberFormats,
            styles.NumberFormats,
            XLPredefinedFormat.NumberFormatIds,
            FirstUserDefinedFormatIndex
        );

        if (numberFormatMap.Count > predefinedNumberFormats.Count)
        {
            this.WriteNumberFormats(xml, numberFormatMap, predefinedNumberFormats);
        }

        // Fonts. Register default format font as font zero. The font zero is used for font name and size.
        Dictionary<XLFontFormatValue, int> firstFontValues = new()
        {
            { styles.DefaultFormat.Font, 0 },
        };
        SequentialMap<int, XLFontFormatValue> fontFormatsMap = SequentialMap<
            int,
            XLFontFormatValue
        >.Create(usedFonts, styles.Fonts, firstFontValues);
        if (fontFormatsMap.Count > 0)
        {
            this.WriteFonts(xml, fontFormatsMap);
        }

        // Fill 0 must be None and fill 1 must be Gray125, that is just an immutable fact of the universe.
        // Excel will ignore fills at 0/1 and will use None/Gray125. Write both fills whether they are used
        // or not.
        AddFillAsUsed(XLFillFormatValue.None);
        AddFillAsUsed(XLFillFormatValue.Gray125);
        Dictionary<XLFillFormatValue, int> firstFillValues = new()
        {
            { XLFillFormatValue.None, 0 },
            { XLFillFormatValue.Gray125, 1 },
        };
        SequentialMap<int, XLFillFormatValue> fillsFormatsMap = SequentialMap<
            int,
            XLFillFormatValue
        >.Create(usedFills, styles.Fills, firstFillValues);
        this.WriteFills(xml, fillsFormatsMap);

        SequentialMap<int, XLBorderFormatValue> borderFormatsMap = SequentialMap<
            int,
            XLBorderFormatValue
        >.Create(usedBorders, styles.Borders);
        if (borderFormatsMap.Count > 0)
        {
            this.WriteBorders(xml, borderFormatsMap);
        }

        SequentialMap<StyleId, XLCellStyleValue> cellStylesMap = new(styles.CellStyles);

        // All cell styles, regardless if they are used or not should be written to the file
        foreach (StyleId styleId in styles.CellStyles.Keys)
        {
            cellStylesMap.Add(styleId);
        }

        cellStylesMap.Sort();
        if (cellStylesMap.Count > 0)
        {
            this.WriteCellStyleXfs(
                xml,
                cellStylesMap,
                numberFormatMap,
                fontFormatsMap,
                fillsFormatsMap,
                borderFormatsMap
            );
        }

        Dictionary<XLCellFormatValue, int> firstCellXfsValues = new()
        {
            { styles.DefaultFormat, 0 },
        };
        SequentialMap<int, XLCellFormatValue> cellXfsMap = SequentialMap<
            int,
            XLCellFormatValue
        >.Create(usedCellFormats, styles.CellFormats, firstCellXfsValues);
        this.WriteCellXfs(
            xml,
            cellXfsMap,
            numberFormatMap,
            fontFormatsMap,
            fillsFormatsMap,
            borderFormatsMap,
            cellStylesMap
        );

        if (cellStylesMap.Count > 0)
        {
            this.WriteCellStyles(xml, cellStylesMap);
        }

        SequentialMap<int, XLDxfValue> dxfMap = SequentialMap<int, XLDxfValue>.Create(
            usedDxf,
            styles.DifferentialFormats
        );
        this.WriteDxfs(xml, dxfMap, numberFormatMap.Count);

        bool hasTableStyles =
            styles.TableStyles.Count > 0
            || styles.PivotStyles.Count > 0
            || styles.DefaultTableStyle is not null
            || styles.DefaultPivotStyle is not null;
        if (hasTableStyles)
        {
            this.WriteTableStyles(xml, dxfMap, styles);
        }

        this.WriteColors(xml, styles);

        this.WriteExtensions(xml);

        xml.WriteEndDocument();

        // Fill the maps used in other parts to determine saved id for a format
        foreach ((int numberFormatId, XLNumberFormat numberFormat) in numberFormatMap.GetActual())
        {
            if (!context.NumberFormatMap.ContainsKey(numberFormat))
            {
                context.NumberFormatMap.Add(numberFormat, numberFormatId);
            }
        }

        foreach ((int fontId, XLFontFormatValue fontFormat) in fontFormatsMap.GetActual())
        {
            if (!context.FontMap.ContainsKey(fontFormat))
            {
                context.FontMap.Add(fontFormat, fontId);
            }
        }

        foreach ((int xfId, XLCellFormatValue format) in cellXfsMap.GetActual())
        {
            if (!context.FormatMap.ContainsKey(format))
            {
                context.FormatMap.Add(format, (uint)xfId);
            }
        }

        foreach ((int dxfId, XLDxfValue dxf) in dxfMap.GetActual())
        {
            if (!context.DxfMap.ContainsKey(dxf))
            {
                context.DxfMap.Add(dxf, (uint)dxfId);
            }
        }

        return;

        void AddFillAsUsed(XLFillFormatValue format)
        {
            if (!styles.Fills.ContainsValue(format))
            {
                styles.AddFillFormat(format);
            }

            usedFills.Add(format);
        }
    }

    private void WriteNumberFormats(
        XmlTreeWriter xml,
        SequentialMap<int, XLNumberFormat> idMap,
        IReadOnlyDictionary<XLNumberFormat, int> predefinedNumberFormats
    )
    {
        xml.WriteStartElement("numFmts", this._ns);
        xml.WriteAttribute("count", idMap.Count - predefinedNumberFormats.Count);

        foreach ((int savedId, XLNumberFormat format) in idMap.GetActual())
        {
            // Do not write the predefined formats, Excel doesn't write that either
            if (predefinedNumberFormats.ContainsKey(format))
            {
                continue;
            }

            this.WriteNumFmt(xml, "numFmt", savedId, format);
        }

        xml.WriteEndElement(); // numFmts
    }

    private void WriteNumFmt(XmlTreeWriter xml, string elementName, int numFmtId, string format)
    {
        xml.WriteStartElement(elementName, this._ns);
        xml.WriteAttribute("numFmtId", numFmtId);
        xml.WriteAttribute("formatCode", format);
        xml.WriteEndElement();
    }

    private void WriteFonts(XmlTreeWriter xml, SequentialMap<int, XLFontFormatValue> idMap)
    {
        xml.WriteStartElement("fonts", this._ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach ((int _, XLFontFormatValue font) in idMap.GetActual())
        {
            this.WriteFont(xml, "font", font);
        }

        xml.WriteEndElement();
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XLFontFormatValue font)
    {
        // MS-OI29500 dictates font elements order.
        xml.WriteStartElement(elementName, this._ns);

        WriteFlag("b", font.Bold);
        WriteFlag("i", font.Italic);
        WriteFlag("strike", font.Strikethrough);
        WriteFlag("condense", font.Condense);
        WriteFlag("extend", font.Extend);
        WriteFlag("outline", font.Outline);
        WriteFlag("shadow", font.Shadow);

        if (font.Underline != XLFontUnderlineValues.None)
        {
            this.WriteFontUnderline(xml, font.Underline);
        }

        if (font.VerticalAlignment != XLFontVerticalTextAlignmentValues.Baseline)
        {
            this.WriteFontVerticalAlignment(xml, font.VerticalAlignment);
        }

        this.WriteFontSize(xml, font.Size);

        if (!font.Color.IsAuto)
        {
            xml.WriteColor("color", this._ns, font.Color);
        }

        this.WriteFontName(xml, font.Name);

        if (font.Family != XLFontFamilyNumberingValues.NotApplicable)
        {
            this.WriteFontFamily(xml, font.Family);
        }

        if (font.Charset != XLFontCharSet.Ansi)
        {
            this.WriteFontCharset(xml, font.Charset);
        }

        if (font.Scheme != XLFontScheme.None)
        {
            this.WriteFontScheme(xml, font.Scheme);
        }

        xml.WriteEndElement();
        return;

        void WriteFlag(string flagName, bool flag)
        {
            if (flag)
            {
                xml.WriteBooleanProperty(flagName, true, this._ns);
            }
        }
    }

    private void WriteFont(XmlTreeWriter xml, string elementName, XLDifferentialFontValue font)
    {
        // MS-OI29500 dictates font elements order.
        xml.WriteStartElement(elementName, this._ns);

        WriteFlag("b", font.Bold);
        WriteFlag("i", font.Italic);
        WriteFlag("strike", font.Strikethrough);
        WriteFlag("condense", font.Condense);
        WriteFlag("extend", font.Extend);
        WriteFlag("outline", font.Outline);
        WriteFlag("shadow", font.Shadow);

        if (font.Underline is { } underline && underline != XLFontUnderlineValues.None)
        {
            this.WriteFontUnderline(xml, underline);
        }

        if (
            font.VerticalAlignment is { } verticalAlignment
            && verticalAlignment != XLFontVerticalTextAlignmentValues.Baseline
        )
        {
            this.WriteFontVerticalAlignment(xml, verticalAlignment);
        }

        if (font.Size is { } size)
        {
            this.WriteFontSize(xml, size);
        }

        if (font.Color is { } color && !color.IsAuto)
        {
            xml.WriteColor("color", this._ns, color);
        }

        if (font.Name is { } name)
        {
            this.WriteFontName(xml, name);
        }

        if (font.Family is { } family && family != XLFontFamilyNumberingValues.NotApplicable)
        {
            this.WriteFontFamily(xml, family);
        }

        if (font.Charset is { } charset && charset != XLFontCharSet.Ansi)
        {
            this.WriteFontCharset(xml, charset);
        }

        if (font.Scheme is { } scheme && scheme != XLFontScheme.None)
        {
            this.WriteFontScheme(xml, scheme);
        }

        xml.WriteEndElement();
        return;

        void WriteFlag(string flagName, bool? flag)
        {
            if (flag == true)
            {
                xml.WriteBooleanProperty(flagName, true, this._ns);
            }
        }
    }

    private void WriteFontUnderline(XmlTreeWriter xml, XLFontUnderlineValues underline)
    {
        xml.WriteStartElement("u", this._ns);
        xml.WriteAttributeDefault("val", underline, XLFontUnderlineValues.Single);
        xml.WriteEndElement();
    }

    private void WriteFontVerticalAlignment(
        XmlTreeWriter xml,
        XLFontVerticalTextAlignmentValues verticalAlignment
    )
    {
        xml.WriteStartElement("vertAlign", this._ns);
        xml.WriteAttribute("val", verticalAlignment);
        xml.WriteEndElement();
    }

    private void WriteFontSize(XmlTreeWriter xml, XLFontSize size)
    {
        xml.WriteStartElement("sz", this._ns);
        xml.WriteAttribute("val", size.Points);
        xml.WriteEndElement();
    }

    private void WriteFontName(XmlTreeWriter xml, XLFontName fontName)
    {
        xml.WriteStartElement("name", this._ns);
        xml.WriteAttribute("val", fontName.Text);
        xml.WriteEndElement();
    }

    private void WriteFontFamily(XmlTreeWriter xml, XLFontFamilyNumberingValues family)
    {
        xml.WriteStartElement("family", this._ns);
        xml.WriteAttribute("val", (int)family);
        xml.WriteEndElement();
    }

    private void WriteFontCharset(XmlTreeWriter xml, XLFontCharSet charset)
    {
        // Charset is stored as an CT_IntProperty
        xml.WriteStartElement("charset", this._ns);
        xml.WriteAttribute("val", (int)charset);
        xml.WriteEndElement();
    }

    private void WriteFontScheme(XmlTreeWriter xml, XLFontScheme scheme)
    {
        xml.WriteStartElement("scheme", this._ns);
        xml.WriteAttribute("val", scheme);
        xml.WriteEndElement();
    }

    private void WriteFills(XmlTreeWriter xml, SequentialMap<int, XLFillFormatValue> idMap)
    {
        xml.WriteStartElement("fills", this._ns);
        xml.WriteAttribute("count", idMap.Count);

        foreach ((int _, XLFillFormatValue fill) in idMap.GetActual())
        {
            this.WriteFill(
                xml,
                "fill",
                fill.Pattern,
                fill.LinearGradient,
                fill.PathGradient,
                false
            );
        }

        xml.WriteEndElement();
    }

    private void WriteFill(
        XmlTreeWriter xml,
        string elementName,
        XLPatternFill? patternFill,
        XLLinearGradientFill? linearGradient,
        XLPathGradientFill? pathGradient,
        bool isDxf
    )
    {
        xml.WriteStartElement(elementName, this._ns);

        // A fill element with no pattern/gradient is a valid state per XML
        if (patternFill is not null)
        {
            xml.WriteStartElement("patternFill", this._ns);
            xml.WriteAttribute("patternType", patternFill.PatternType);

            XLColor patternColor = patternFill.PatternColor;
            XLColor bgColor = patternFill.BackgroundColor;

            // Fix solid pattern discrepancy. The GUI shows solid fill color in the background
            // color picker, so it would be expected that it is stored in bgColor.
            // * internal structures store solid fill color in the 'IXLFill.BackgroundColor'
            // * For cell format, the 'patternFill' element stores it in the *fgColor*, not bgColor
            // * For dxf, the 'patternFill' stores it in the *bgColor*
            if (!isDxf && patternFill.PatternType == XLFillPatternValues.Solid)
            {
                (patternColor, bgColor) = (bgColor, patternColor);
            }

            if (patternColor.HasValue)
            {
                xml.WriteColor("fgColor", this._ns, patternColor);
            }

            if (bgColor.HasValue)
            {
                xml.WriteColor("bgColor", this._ns, bgColor);
            }

            xml.WriteEndElement();
        }
        else if (linearGradient is not null)
        {
            // Linear is the default type, so no need to write it
            xml.WriteStartElement("gradientFill", this._ns);
            xml.WriteAttributeDefault("degree", linearGradient.Degrees, 0);

            WriteStops(linearGradient.Stops);

            xml.WriteEndElement();
        }
        else if (pathGradient is not null)
        {
            xml.WriteStartElement("gradientFill", this._ns);
            xml.WriteAttribute("type", "path");
            xml.WriteAttributeDefault("left", pathGradient.InnerLeft.Value, 0);
            xml.WriteAttributeDefault("right", pathGradient.InnerRight.Value, 0);
            xml.WriteAttributeDefault("top", pathGradient.InnerTop.Value, 0);
            xml.WriteAttributeDefault("bottom", pathGradient.InnerBottom.Value, 0);

            WriteStops(pathGradient.Stops);

            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        void WriteStops(IReadOnlyDictionary<FractionOfOne, XLColor> stops)
        {
            // Excel doesn't care about stop order by positions, but sort anyway
            foreach ((FractionOfOne position, XLColor color) in stops.OrderBy(x => x.Key.Value))
            {
                xml.WriteStartElement("stop", this._ns);
                xml.WriteAttribute("position", position.Value);
                xml.WriteColor("color", this._ns, color);
                xml.WriteEndElement();
            }
        }
    }

    private void WriteBorders(XmlTreeWriter xml, SequentialMap<int, XLBorderFormatValue> idMap)
    {
        xml.WriteStartElement("borders", this._ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach ((int _, XLBorderFormatValue border) in idMap.GetActual())
        {
            this.WriteBorder(xml, "border", border);
        }

        xml.WriteEndElement();
    }

    private void WriteBorder(XmlTreeWriter xml, string elementName, XLBorderFormatValue border)
    {
        xml.WriteStartElement(elementName, this._ns);
        xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
        xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);

        // ISO should be "start"+"end", but Excel uses "left"+"right"
        this.WriteBorderPr(xml, "left", border.Left);
        this.WriteBorderPr(xml, "right", border.Right);
        this.WriteBorderPr(xml, "top", border.Top);
        this.WriteBorderPr(xml, "bottom", border.Bottom);
        this.WriteBorderPr(xml, "diagonal", border.Diagonal);

        xml.WriteEndElement();
    }

    private void WriteBorder(
        XmlTreeWriter xml,
        string elementName,
        XLDifferentialBorderValue border
    )
    {
        xml.WriteStartElement(elementName, this._ns);
        xml.WriteAttributeDefault("diagonalUp", border.DiagonalUp, false);
        xml.WriteAttributeDefault("diagonalDown", border.DiagonalDown, false);

        // Outline has no meaning for cell styles, it is only for dxf - tables and such.
        xml.WriteAttributeDefault("outline", border.Outline, true);

        // ISO should be "start"+"end", but Excel uses "left"+"right"
        if (border.Left is { } left)
        {
            this.WriteBorderPr(xml, "left", left);
        }

        if (border.Right is { } right)
        {
            this.WriteBorderPr(xml, "right", right);
        }

        if (border.Top is { } top)
        {
            this.WriteBorderPr(xml, "top", top);
        }

        if (border.Bottom is { } bottom)
        {
            this.WriteBorderPr(xml, "bottom", bottom);
        }

        if (border.Diagonal is { } diagonal)
        {
            this.WriteBorderPr(xml, "diagonal", diagonal);
        }

        if (border.Vertical is { } vertical)
        {
            this.WriteBorderPr(xml, "vertical", vertical);
        }

        if (border.Horizontal is { } horizontal)
        {
            this.WriteBorderPr(xml, "horizontal", horizontal);
        }

        xml.WriteEndElement();
    }

    private void WriteBorderPr(XmlTreeWriter xml, string name, XLBorderLine borderLine)
    {
        xml.WriteStartElement(name, this._ns);
        xml.WriteAttributeDefault("style", borderLine.Style, XLBorderStyleValues.None);
        if (borderLine.Style != XLBorderStyleValues.None)
        {
            // Color for border is always written and default value is automatic color.
            xml.WriteColor("color", this._ns, borderLine.Color);
        }

        xml.WriteEndElement();
    }

    private void WriteCellStyleXfs(
        XmlTreeWriter xml,
        SequentialMap<StyleId, XLCellStyleValue> cellStylesMap,
        SequentialMap<int, XLNumberFormat> numFmtIdMap,
        SequentialMap<int, XLFontFormatValue> fontIdMap,
        SequentialMap<int, XLFillFormatValue> fillIdMap,
        SequentialMap<int, XLBorderFormatValue> borderIdMap
    )
    {
        xml.WriteStartElement("cellStyleXfs", this._ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach ((int _, XLCellStyleValue cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("xf", this._ns);
            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellStyle.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellStyle.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellStyle.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellStyle.Border));

            // cellStyleXf doesn't use quote, pivot button or xfId -> skip those attributes
            xml.WriteAttributeDefault(
                "applyNumberFormat",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.NumberFormat),
                true
            );
            xml.WriteAttributeDefault(
                "applyFont",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Font),
                true
            );
            xml.WriteAttributeDefault(
                "applyFill",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Fill),
                true
            );
            xml.WriteAttributeDefault(
                "applyBorder",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Border),
                true
            );
            xml.WriteAttributeDefault(
                "applyAlignment",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Alignment),
                true
            );
            xml.WriteAttributeDefault(
                "applyProtection",
                cellStyle.IncludedComponents.HasFlag(CellFormatComponents.Protection),
                true
            );

            if (cellStyle.Alignment is { } alignment)
            {
                this.WriteAlignment(xml, "alignment", alignment);
            }

            if (cellStyle.Protection is { } protection)
            {
                this.WriteProtection(xml, "protection", protection);
            }

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteCellXfs(
        XmlTreeWriter xml,
        SequentialMap<int, XLCellFormatValue> idMap,
        SequentialMap<int, XLNumberFormat> numFmtIdMap,
        SequentialMap<int, XLFontFormatValue> fontIdMap,
        SequentialMap<int, XLFillFormatValue> fillIdMap,
        SequentialMap<int, XLBorderFormatValue> borderIdMap,
        SequentialMap<StyleId, XLCellStyleValue> cellStyleIdMap
    )
    {
        xml.WriteStartElement("cellXfs", this._ns);
        xml.WriteAttribute("count", idMap.Count);
        foreach ((int _, XLCellFormatValue cellXf) in idMap.GetActual())
        {
            xml.WriteStartElement("xf", this._ns);

            xml.WriteAttributeOptional("numFmtId", numFmtIdMap.GetSavedId(cellXf.NumberFormat));
            xml.WriteAttributeOptional("fontId", fontIdMap.GetSavedId(cellXf.Font));
            xml.WriteAttributeOptional("fillId", fillIdMap.GetSavedId(cellXf.Fill));
            xml.WriteAttributeOptional("borderId", borderIdMap.GetSavedId(cellXf.Border));

            if (cellXf.CellStyleId is not null)
            {
                xml.WriteAttribute("xfId", cellStyleIdMap.GetSavedId(cellXf.CellStyleId.Value));
            }

            xml.WriteAttributeDefault("quotePrefix", cellXf.IncludeQuotePrefix, false);
            xml.WriteAttributeDefault("pivotButton", cellXf.PivotButton, false);
            xml.WriteAttributeDefault(
                "applyNumberFormat",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.NumberFormat),
                false
            );
            xml.WriteAttributeDefault(
                "applyFont",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.Font),
                false
            );
            xml.WriteAttributeDefault(
                "applyFill",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.Fill),
                false
            );
            xml.WriteAttributeDefault(
                "applyBorder",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.Border),
                false
            );
            xml.WriteAttributeDefault(
                "applyAlignment",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.Alignment),
                false
            );
            xml.WriteAttributeDefault(
                "applyProtection",
                cellXf.CustomFormat.HasFlag(CellFormatComponents.Protection),
                false
            );

            if (cellXf.Alignment is { } alignment)
            {
                this.WriteAlignment(xml, "alignment", alignment);
            }

            if (cellXf.Protection is { } protection)
            {
                this.WriteProtection(xml, "protection", protection);
            }

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteAlignment(
        XmlTreeWriter xml,
        string elementName,
        XLAlignmentFormatValue alignment
    )
    {
        if (alignment == XLAlignmentFormatValue.Default)
        {
            return;
        }

        xml.WriteStartElement(elementName, this._ns);
        xml.WriteAttributeDefault(
            "horizontal",
            alignment.Horizontal,
            XLAlignmentFormatValue.Default.Horizontal
        );
        xml.WriteAttributeDefault(
            "vertical",
            alignment.Vertical,
            XLAlignmentFormatValue.Default.Vertical
        );
        xml.WriteAttributeDefault(
            "textRotation",
            alignment.TextRotation.GetIso(),
            XLAlignmentFormatValue.Default.TextRotation.Value
        );
        xml.WriteAttributeDefault(
            "wrapText",
            alignment.WrapText,
            XLAlignmentFormatValue.Default.WrapText
        );
        xml.WriteAttributeDefault(
            "indent",
            alignment.Indent,
            XLAlignmentFormatValue.Default.Indent
        );
        xml.WriteAttributeDefault(
            "relativeIndent",
            alignment.RelativeIndent,
            XLAlignmentFormatValue.Default.RelativeIndent
        );
        xml.WriteAttributeDefault(
            "justifyLastLine",
            alignment.JustifyLastLine,
            XLAlignmentFormatValue.Default.JustifyLastLine
        );
        xml.WriteAttributeDefault(
            "shrinkToFit",
            alignment.ShrinkToFit,
            XLAlignmentFormatValue.Default.ShrinkToFit
        );
        xml.WriteAttributeDefault(
            "readingOrder",
            (uint)alignment.ReadingOrder,
            (uint)XLAlignmentFormatValue.Default.ReadingOrder
        );
        xml.WriteEndElement();
    }

    private void WriteAlignment(
        XmlTreeWriter xml,
        string elementName,
        XLDifferentialAlignmentValue alignment
    )
    {
        // Alignment is kind of wonky. Current Excel doesn't even support it in a DXF dialogue and doesn't always work correctly.
        xml.WriteStartElement(elementName, this._ns);
        if (alignment.Horizontal is { } horizontal)
        {
            xml.WriteAttribute("horizontal", horizontal);
        }

        if (alignment.Vertical is { } vertical)
        {
            xml.WriteAttribute("vertical", vertical);
        }

        if (alignment.TextRotation is { } textRotation)
        {
            xml.WriteAttribute("textRotation", textRotation.GetIso());
        }

        if (alignment.WrapText is { } wrapText)
        {
            xml.WriteAttribute("wrapText", wrapText);
        }

        if (alignment.Indent is { } indent)
        {
            xml.WriteAttribute("indent", indent);
        }

        if (alignment.RelativeIndent is { } relativeIndent)
        {
            xml.WriteAttribute("relativeIndent", relativeIndent);
        }

        if (alignment.JustifyLastLine is { } justifyLastLine)
        {
            xml.WriteAttribute("justifyLastLine", justifyLastLine);
        }

        if (alignment.ShrinkToFit is { } shrinkToFit)
        {
            xml.WriteAttribute("shrinkToFit", shrinkToFit);
        }

        if (alignment.ReadingOrder is { } readingOrder)
        {
            xml.WriteAttribute("readingOrder", readingOrder);
        }

        xml.WriteEndElement();
    }

    private void WriteProtection(
        XmlTreeWriter xml,
        string elementName,
        XLProtectionFormatValue protection
    )
    {
        if (protection.Locked && !protection.Hidden)
        {
            return;
        }

        xml.WriteStartElement(elementName, this._ns);
        xml.WriteAttributeDefault("locked", protection.Locked, true);
        xml.WriteAttributeDefault("hidden", protection.Hidden, false);
        xml.WriteEndElement();
    }

    private void WriteProtection(
        XmlTreeWriter xml,
        string elementName,
        XLDifferentialProtectionValue protection
    )
    {
        xml.WriteStartElement(elementName, this._ns);

        if (protection.Locked is { } locked)
        {
            xml.WriteAttribute("locked", locked);
        }

        if (protection.Hidden is { } hidden)
        {
            xml.WriteAttribute("hidden", hidden);
        }

        xml.WriteEndElement();
    }

    private void WriteCellStyles(
        XmlTreeWriter xml,
        SequentialMap<StyleId, XLCellStyleValue> cellStylesMap
    )
    {
        xml.WriteStartElement("cellStyles", this._ns);
        xml.WriteAttribute("count", cellStylesMap.Count);

        // Collection must have at least one element
        foreach ((int mappedStyleId, XLCellStyleValue cellStyle) in cellStylesMap.GetActual())
        {
            xml.WriteStartElement("cellStyle", this._ns);

            // Name is technically optional and Excel will generate one if missing, but we ensure the name always exist
            xml.WriteAttribute("name", cellStyle.Name);
            xml.WriteAttribute("xfId", mappedStyleId);
            if (cellStyle.BuiltInStyle is { } builtInStyle)
            {
                if (
                    builtInStyle
                    is >= BuiltInStyleValues.RowLevel1
                        and <= BuiltInStyleValues.RowLevel7
                )
                {
                    xml.WriteAttributeOptional("builtinId", 1);
                    xml.WriteAttribute("iLevel", BuiltInStyleValues.RowLevel1 - builtInStyle + 1);
                }
                else if (
                    builtInStyle
                    is >= BuiltInStyleValues.ColumnLevel1
                        and <= BuiltInStyleValues.ColumnLevel7
                )
                {
                    xml.WriteAttributeOptional("builtinId", 2);
                    xml.WriteAttribute(
                        "iLevel",
                        BuiltInStyleValues.ColumnLevel1 - builtInStyle + 1
                    );
                }
                else
                {
                    xml.WriteAttributeOptional("builtinId", (int?)cellStyle.BuiltInStyle);
                }
            }

            // Hidden + flag are optional per schema, but basically it's a bool with default
            xml.WriteAttributeDefault("hidden", cellStyle.Hidden, false);
            xml.WriteAttributeDefault("customBuiltin", cellStyle.BuiltInStyle is not null, true);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteDxfs(
        XmlTreeWriter xml,
        SequentialMap<int, XLDxfValue> differentialFormats,
        int lastNumFmtId
    )
    {
        xml.WriteStartElement("dxfs", this._ns);
        xml.WriteAttribute("count", differentialFormats.Count);
        foreach ((int _, XLDxfValue dxf) in differentialFormats.GetActual())
        {
            xml.WriteStartElement("dxf", this._ns);

            if (dxf.Font != XLDifferentialFontValue.Empty)
            {
                this.WriteFont(xml, "font", dxf.Font);
            }

            if (dxf.NumberFormat is { } numberFormat)
            {
                // numFmtId doesn't matter in dxf, but keep them unique (Excel-like behavior)
                this.WriteNumFmt(xml, "numFmt", ++lastNumFmtId, numberFormat);
            }

            if (dxf.Fill != XLDifferentialFillValue.Empty)
            {
                this.WriteFill(
                    xml,
                    "fill",
                    dxf.Fill.Pattern,
                    dxf.Fill.LinearGradient,
                    dxf.Fill.PathGradient,
                    true
                );
            }

            if (dxf.Alignment != XLDifferentialAlignmentValue.Empty)
            {
                this.WriteAlignment(xml, "alignment", dxf.Alignment);
            }

            if (dxf.Border != XLDifferentialBorderValue.Empty)
            {
                this.WriteBorder(xml, "border", dxf.Border);
            }

            if (dxf.Protection != XLDifferentialProtectionValue.Empty)
            {
                this.WriteProtection(xml, "protection", dxf.Protection);
            }

            // TODO: extLst
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
    }

    private void WriteTableStyles(
        XmlTreeWriter xml,
        SequentialMap<int, XLDxfValue> dxfMap,
        XLWorkbookStyles styles
    )
    {
        List<string> allStyleNames =
        [
            .. styles
                .TableStyles.Keys.Concat(styles.PivotStyles.Keys)
                .Distinct(XlsxSharp.XLHelper.NameComparer),
        ];
        allStyleNames.Sort();

        xml.WriteStartElement("tableStyles", this._ns);
        xml.WriteAttribute("count", allStyleNames.Count);
        if (styles.DefaultTableStyle is { } defaultTableStyle)
        {
            xml.WriteAttribute("defaultTableStyle", defaultTableStyle);
        }

        if (styles.DefaultPivotStyle is { } defaultPivotStyle)
        {
            xml.WriteAttribute("defaultPivotStyle", defaultPivotStyle);
        }

        foreach (string styleName in allStyleNames)
        {
            this.WriteTableStyle(xml, styleName, dxfMap, styles);
        }

        xml.WriteEndElement();
    }

    private void WriteTableStyle(
        XmlTreeWriter xml,
        string styleName,
        SequentialMap<int, XLDxfValue> dxfMap,
        XLWorkbookStyles styles
    )
    {
        xml.WriteStartElement("tableStyle", this._ns);
        xml.WriteAttribute("name", styleName);

        bool hasPivotStyle = styles.PivotStyles.TryGetValue(
            styleName,
            out XLPivotTableStyle? pivotStyle
        );
        xml.WriteAttributeDefault("pivot", hasPivotStyle, true);

        bool hasTableStyle = styles.TableStyles.TryGetValue(
            styleName,
            out XLTableTheme? tableStyle
        );
        xml.WriteAttributeDefault("table", hasTableStyle, true);

        List<(string Type, int Size, int DxfId)> styledRegions = new(TableRegionsMap.Count);
        foreach (
            (string type, TableRegion? tableRegion, PivotRegion? pivotRegion) in TableRegionsMap
        )
        {
            (int Size, XLDxfValue Dxf)? tableRegionStyle = TryGetTableRegion(
                tableStyle,
                tableRegion
            );
            (int Size, XLDxfValue Dxf)? pivotRegionStyle = TryGetPivotRegion(
                pivotStyle,
                pivotRegion
            );
            if (
                tableRegionStyle is var (tableSize, tableDxf)
                && pivotRegionStyle is var (pivotSize, pivotDxf)
                && (tableSize != pivotSize || tableDxf != pivotDxf)
            )
            {
                // This should never happen. The table/pivot shared style have same band size and
                // dxf on load and we don't provide API to modify table/pivot styles.
                // Sidenote: Excel GUI will refuse to create table style that conflicts with a pivot
                // style and vice versa. It will show an alert 'This style name already exists.'
                throw new InvalidOperationException(
                    $"Table and pivot table style '{styleName}' that has different formatting for {tableRegion}/{pivotRegion}."
                );
            }

            if ((tableRegionStyle ?? pivotRegionStyle) is var (bandSize, dxf))
            {
                styledRegions.Add((type, bandSize, dxfMap.GetSavedId(dxf)));
            }
        }

        xml.WriteAttribute("count", styledRegions.Count);
        foreach ((string type, int bandSize, int dxfId) in styledRegions)
        {
            xml.WriteStartElement("tableStyleElement", this._ns);
            xml.WriteAttribute("type", type);
            xml.WriteAttributeDefault("size", bandSize, 1);
            xml.WriteAttribute("dxfId", dxfId);
            xml.WriteEndElement();
        }

        xml.WriteEndElement();
        return;

        static (int Size, XLDxfValue Dxf)? TryGetTableRegion(
            XLTableTheme? tableStyle,
            TableRegion? tableRegion
        )
        {
            if (tableStyle is null || tableRegion is null)
            {
                return null;
            }

            if (!tableStyle.RegionFormats.TryGetValue(tableRegion.Value, out XLDxfValue? dxf))
            {
                return null;
            }

            int bandSize = tableRegion switch
            {
                TableRegion.FirstRowStripe => tableStyle.RowStripe1BandSize,
                TableRegion.SecondRowStripe => tableStyle.RowStripe2BandSize,
                TableRegion.FirstColumnStripe => tableStyle.ColumnStripe1BandSize,
                TableRegion.SecondColumnStripe => tableStyle.ColumnStripe2BandSize,
                _ => 1,
            };
            return (bandSize, dxf);
        }

        static (int Size, XLDxfValue Dxf)? TryGetPivotRegion(
            XLPivotTableStyle? pivotStyle,
            PivotRegion? region
        )
        {
            if (pivotStyle is null || region is null)
            {
                return null;
            }

            if (!pivotStyle.RegionFormats.TryGetValue(region.Value, out XLDxfValue? dxf))
            {
                return null;
            }

            int bandSize = region switch
            {
                PivotRegion.FirstRowStripe => pivotStyle.RowStripe1BandSize,
                PivotRegion.SecondRowStripe => pivotStyle.RowStripe2BandSize,
                PivotRegion.FirstColumnStripe => pivotStyle.ColumnStripe1BandSize,
                PivotRegion.SecondColumnStripe => pivotStyle.ColumnStripe2BandSize,
                _ => 1,
            };
            return (bandSize, dxf);
        }
    }

    private void WriteColors(XmlTreeWriter xml, XLWorkbookStyles styles)
    {
        bool hasMruColors = styles.MruColors.Count > 0;
        IReadOnlyList<uint>? indexedColors = styles.IndexedColorsArgb;
        bool hasIndexColors = indexedColors is { Count: > 0 };
        if (!hasMruColors && !hasIndexColors)
        {
            return;
        }

        xml.WriteStartElement("colors", this._ns);

        if (hasIndexColors)
        {
            xml.WriteStartElement("indexedColors", this._ns);
            foreach (uint indexedColor in indexedColors!)
            {
                xml.WriteStartElement("rgbColor", this._ns);
                xml.WriteAttributeHex("rgb", indexedColor);
                xml.WriteEndElement();
            }

            xml.WriteEndElement();
        }

        if (hasMruColors)
        {
            xml.WriteStartElement("mruColors", this._ns);
            foreach (XLColor mruColor in styles.MruColors)
            {
                xml.WriteColor("color", this._ns, mruColor);
            }

            xml.WriteEndElement();
        }

        xml.WriteEndElement(); // colors
    }

    private void WriteExtensions(XmlTreeWriter xml)
    {
        xml.WriteStartElement("extLst", this._ns);
        this.WriteSlicerStyles(xml);
        this.WriteTimelineStyles(xml);
        xml.WriteEndElement();
    }

    private void WriteSlicerStyles(XmlTreeWriter xml)
    {
        // TODO Styles: Represent and write back, this only writes default created by Excel
        xml.WriteStartExtension(
            "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
            this._ns,
            "x14",
            X14Main2009SsNs
        );
        xml.WriteStartElement("slicerStyles", X14Main2009SsNs);
        xml.WriteAttribute("defaultSlicerStyle", "SlicerStyleLight1");
        xml.WriteEndElement();
        xml.WriteEndElement();

        // dxfs for slicer styles: 46F421CA-312F-682F-3DD2-61675219B42D
    }

    private void WriteTimelineStyles(XmlTreeWriter xml)
    {
        // TODO Styles: Represent and write back, this only writes default created by Excel
        xml.WriteStartExtension(
            "{9260A510-F301-46a8-8635-F512D64BE5F5}",
            this._ns,
            "x15",
            X15Main2010SsNs
        );
        xml.WriteStartElement("timelineStyles", X15Main2010SsNs);
        xml.WriteAttribute("defaultTimelineStyle", "TimeSlicerStyleLight1");
        xml.WriteEndElement();
        xml.WriteEndElement();

        // dxfs for timeline styles: A0A4C193-F2C1-4fcb-8827-314CF55A85BB
    }
}

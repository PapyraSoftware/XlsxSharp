using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.CalcEngine;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.DataValidation;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Excel.PageSetup;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.Parser;
using XlsxSharp.Utils;
using Formula = DocumentFormat.OpenXml.Spreadsheet.Formula;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XlsxSharp.Excel.IO;

#nullable disable

internal class WorksheetPartReader
{
    private static readonly string[] DateCellFormats =
    [
        "yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", // Format accepted by OpenXML SDK
        "yyyy-MM-ddTHH:mm",
        "yyyy-MM-dd", // Formats accepted by Excel.
    ];

    private readonly Dictionary<uint, string> _sharedFormulasR1C1 = new();

    /// <summary>
    /// Row number of last read <c>row</c> element.
    /// </summary>
    private int _lastRow;
    private int _lastColumnNumber;

    internal void LoadWorksheet(
        XLWorksheet ws,
        WorksheetPart worksheetPart,
        SharedStringItem[] sharedStrings,
        LoadContext context
    )
    {
        XElement pageSetupProperties = null;

        this._lastRow = 0;

        using (OpenXmlPartReader reader = new(worksheetPart))
        {
            Type[] ignoredElements =
            [
                typeof(CustomSheetViews), // Custom sheet views contain its own auto filter data, and more, which should be ignored for now
            ];

            while (reader.Read())
            {
                while (ignoredElements.Contains(reader.ElementType))
                {
                    reader.ReadNextSibling();
                }

                if (reader.ElementType == typeof(SheetFormatProperties))
                {
                    LoadSheetFormatProperties(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(SheetViews))
                {
                    LoadSheetViews(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(MergeCells))
                {
                    LoadMergeCells(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(Columns))
                {
                    LoadColumns(ws, AsXElement(reader.LoadCurrentElement()));
                }
                else if (reader.ElementType == typeof(Row))
                {
                    this.LoadRow(ws, sharedStrings, reader);
                }
                else if (reader.ElementType == typeof(AutoFilter))
                {
                    AutoFilterReader.LoadAutoFilter((AutoFilter)reader.LoadCurrentElement(), ws);
                }
                else if (reader.ElementType == typeof(SheetProtection))
                {
                    LoadSheetProtection(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(DataValidations))
                {
                    LoadDataValidations(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(ConditionalFormatting))
                {
                    LoadConditionalFormatting(AsXElement(reader.LoadCurrentElement()), ws, context);
                }
                else if (reader.ElementType == typeof(Hyperlinks))
                {
                    LoadHyperlinks(AsXElement(reader.LoadCurrentElement()), worksheetPart, ws);
                }
                else if (reader.ElementType == typeof(PrintOptions))
                {
                    LoadPrintOptions(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(PageMargins))
                {
                    LoadPageMargins(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.PageSetup))
                {
                    LoadPageSetup(AsXElement(reader.LoadCurrentElement()), ws, pageSetupProperties);
                }
                else if (reader.ElementType == typeof(HeaderFooter))
                {
                    LoadHeaderFooter(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(SheetProperties))
                {
                    LoadSheetProperties(
                        AsXElement(reader.LoadCurrentElement()),
                        ws,
                        out pageSetupProperties
                    );
                }
                else if (reader.ElementType == typeof(RowBreaks))
                {
                    LoadRowBreaks(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(ColumnBreaks))
                {
                    LoadColumnBreaks(AsXElement(reader.LoadCurrentElement()), ws);
                }
                else if (reader.ElementType == typeof(WorksheetExtensionList))
                {
                    LoadExtensions((WorksheetExtensionList)reader.LoadCurrentElement(), ws);
                }
                else if (reader.ElementType == typeof(LegacyDrawing))
                {
                    ws.LegacyDrawingId = AsXElement(reader.LoadCurrentElement())
                        .Attribute(SpreadsheetXml.Rel + "id")
                        ?.Value;
                }
            }
            reader.Close();
        }
    }

    private static void LoadSheetFormatProperties(XElement sheetFormatProperties, XLWorksheet ws)
    {
        if (sheetFormatProperties is null)
        {
            return;
        }

        if (SpreadsheetXml.Double(sheetFormatProperties, "defaultRowHeight") is { } rowHeight)
        {
            ws.RowHeight = rowHeight;
        }

        ws.RowHeightChanged = SpreadsheetXml.Bool(sheetFormatProperties, "customHeight") ?? false;

        if (SpreadsheetXml.Double(sheetFormatProperties, "defaultColWidth") is { } columnWidth)
        {
            ws.ColumnWidth = XlsxSharp.XLHelper.ConvertWidthToNoC(
                columnWidth,
                ws.Style.Font,
                ws.Workbook
            );
        }
        else if (SpreadsheetXml.UInt(sheetFormatProperties, "baseColWidth") is { } baseWidth)
        {
            ws.ColumnWidth = XlsxSharp.XLHelper.CalculateColumnWidth(
                baseWidth,
                ws.Style.Font,
                ws.Workbook
            );
        }
    }

    private static void LoadMergeCells(XElement mergedCells, XLWorksheet ws)
    {
        if (mergedCells is null)
        {
            return;
        }

        foreach (XElement mergeCell in mergedCells.Elements(SpreadsheetXml.Main + "mergeCell"))
        {
            ws.Range(SpreadsheetXml.String(mergeCell, "ref")).Merge(false);
        }
    }

    private static void LoadSheetProperties(
        XElement sheetProperty,
        XLWorksheet ws,
        out XElement pageSetupProperties
    )
    {
        pageSetupProperties = null;
        if (sheetProperty is null)
        {
            return;
        }

        if (sheetProperty.Element(SpreadsheetXml.Main + "tabColor") is { } tabColor)
        {
            ws.TabColor = SpreadsheetXml.ReadColor(tabColor);
        }

        if (sheetProperty.Element(SpreadsheetXml.Main + "outlinePr") is { } outline)
        {
            if (SpreadsheetXml.Bool(outline, "summaryBelow") is { } summaryBelow)
            {
                ws.Outline.SummaryVLocation = summaryBelow
                    ? XLOutlineSummaryVLocation.Bottom
                    : XLOutlineSummaryVLocation.Top;
            }

            if (SpreadsheetXml.Bool(outline, "summaryRight") is { } summaryRight)
            {
                ws.Outline.SummaryHLocation = summaryRight
                    ? XLOutlineSummaryHLocation.Right
                    : XLOutlineSummaryHLocation.Left;
            }
        }

        // The fitToPage flag lives here, but the page counts it turns on are on pageSetup, which
        // is read further down the part - so hand it to the caller to thread through.
        pageSetupProperties = sheetProperty.Element(SpreadsheetXml.Main + "pageSetUpPr");
    }

    private static void LoadColumns(XLWorksheet ws, XElement columns)
    {
        if (columns is null)
        {
            return;
        }

        XElement[] cols = [.. columns.Elements(SpreadsheetXml.Main + "col")];

        XElement wsDefaultColumn = cols.FirstOrDefault(c =>
            SpreadsheetXml.UInt(c, "max") == XlsxSharp.XLHelper.MaxColumnNumber
        );

        if (SpreadsheetXml.Double(wsDefaultColumn, "width") is { } defaultWidth)
        {
            ws.ColumnWidth = defaultWidth - XLConstants.ColumnWidthOffset;
        }

        // Sheet doesn't have a format, only column spans have format. When whole sheet is selected
        // to change format, Excel will mark all cols spans as having a particular format. Format
        // is considered a sheet format when all columns have a format and it's in the last column.
        (uint? MinColumn, uint? MaxColumn, uint XfId)[] colSpanFormats =
        [
            .. cols.Select(c =>
                (
                    MinColumn: SpreadsheetXml.UInt(c, "min"),
                    MaxColumn: SpreadsheetXml.UInt(c, "max"),
                    XfId: SpreadsheetXml.UInt(c, "style") ?? 0
                )
            ),
        ];
        bool allColsHaveFormat =
            colSpanFormats.Sum(x => x.MaxColumn - x.MinColumn + 1)
            == XlsxSharp.XLHelper.MaxColumnNumber;
        if (allColsHaveFormat)
        {
            uint lastColumnXfId = colSpanFormats
                .Single(x => x.MaxColumn == XlsxSharp.XLHelper.MaxColumnNumber)
                .XfId;
            ApplyStyle(ws, checked((int)lastColumnXfId), ws.Workbook.Styles);
        }

        foreach (XElement col in cols)
        {
            uint? max = SpreadsheetXml.UInt(col, "max");
            if (max == XlsxSharp.XLHelper.MaxColumnNumber)
            {
                continue;
            }

            XLColumns xlColumns = (XLColumns)
                ws.Columns(checked((int)SpreadsheetXml.UInt(col, "min")), checked((int)max));
            xlColumns.Width = SpreadsheetXml.Double(col, "width") is { } width
                ? width - XLConstants.ColumnWidthOffset
                : ws.ColumnWidth;

            if (SpreadsheetXml.Bool(col, "hidden") ?? false)
            {
                xlColumns.Hide();
            }

            if (SpreadsheetXml.Bool(col, "collapsed") ?? false)
            {
                xlColumns.CollapseOnly();
            }

            if (SpreadsheetXml.UInt(col, "outlineLevel") is { } outlineLevel)
            {
                xlColumns.ForEach(c => c.OutlineLevel = checked((int)outlineLevel));
            }

            if (SpreadsheetXml.UInt(col, "style") is { } styleIndex)
            {
                ApplyStyle(xlColumns, checked((int)styleIndex), ws.Workbook.Styles);
            }
        }
    }

    private void LoadRow(XLWorksheet ws, SharedStringItem[] sharedStrings, OpenXmlPartReader reader)
    {
        Debug.Assert(reader.LocalName == "row");

        ReadOnlyCollection<OpenXmlAttribute> attributes = reader.Attributes;
        string rowIndexAttr = attributes.GetAttribute("r");

        // Row number is an optional attribute. If not specified, it should be a next row from the last read row.
        int rowIndex = string.IsNullOrEmpty(rowIndexAttr)
            ? ++this._lastRow
            : int.Parse(rowIndexAttr);
        this._lastRow = rowIndex;

        XLRow xlRow = ws.Row(rowIndex, false);

        double? height = attributes.GetDoubleAttribute("ht");
        if (height is not null)
        {
            xlRow.Height = height.Value;
        }
        else
        {
            xlRow.Loading = true;
            xlRow.Height = ws.RowHeight;
            xlRow.Loading = false;
        }

        double? dyDescent = attributes.GetDoubleAttribute("dyDescent", OpenXmlConst.X14Ac2009SsNs);
        if (dyDescent is not null)
        {
            xlRow.DyDescent = dyDescent.Value;
        }

        bool hidden = attributes.GetBoolAttribute("hidden", false);
        if (hidden)
        {
            xlRow.Hide();
        }

        bool collapsed = attributes.GetBoolAttribute("collapsed", false);
        if (collapsed)
        {
            xlRow.Collapsed = true;
        }

        int? outlineLevel = attributes.GetIntAttribute("outlineLevel");
        if (outlineLevel is not null && outlineLevel.Value > 0)
        {
            xlRow.OutlineLevel = outlineLevel.Value;
        }

        bool showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
        {
            xlRow.ShowPhonetic = true;
        }

        bool customFormat = attributes.GetBoolAttribute("customFormat", false);
        if (customFormat)
        {
            int? styleIndex = attributes.GetIntAttribute("s");
            if (styleIndex is not null)
            {
                ApplyStyle(xlRow, styleIndex.Value, ws.Workbook.Styles);
            }
        }

        this._lastColumnNumber = 0;

        // Move from the start element of 'row' forward. We can get cell, extList or end of row.
        reader.MoveAhead();

        while (reader.IsStartElement("c"))
        {
            this.LoadCell(sharedStrings, ws, reader, rowIndex);

            // Move from end element of 'cell' either to next cell, extList start or end of row.
            reader.MoveAhead();
        }

        // In theory, row can also contain extList, just skip them.
        while (reader.IsStartElement("extLst"))
        {
            reader.Skip();
        }
    }

    private void LoadCell(
        SharedStringItem[] sharedStrings,
        XLWorksheet ws,
        OpenXmlPartReader reader,
        int rowIndex
    )
    {
        Debug.Assert(reader.LocalName == "c" && reader.IsStartElement);

        ReadOnlyCollection<OpenXmlAttribute> attributes = reader.Attributes;

        Point cellAddress =
            attributes.GetCellRefAttribute("r") ?? new Point(rowIndex, this._lastColumnNumber + 1);
        this._lastColumnNumber = cellAddress.Column;

        CellValues dataType = attributes.GetAttribute("t") switch
        {
            "b" => CellValues.Boolean,
            "n" => CellValues.Number,
            "e" => CellValues.Error,
            "s" => CellValues.SharedString,
            "str" => CellValues.String,
            "inlineStr" => CellValues.InlineString,
            "d" => CellValues.Date,
            null => CellValues.Number,
            _ => throw new FormatException($"Unknown cell type."),
        };

        XLCell xlCell = ws.Cell(cellAddress.Row, cellAddress.Column);

        int xfId = attributes.GetIntAttribute("s") ?? 0;
        XLCellFormatValue cellFormat = ws.Workbook.Styles.CellFormats[xfId];
        xlCell.FormatValue = cellFormat;

        bool showPhonetic = attributes.GetBoolAttribute("ph", false);
        if (showPhonetic)
        {
            xlCell.ShowPhonetic = true;
        }

        uint? cellMetaIndex = attributes.GetUintAttribute("cm");
        if (cellMetaIndex is not null)
        {
            xlCell.CellMetaIndex = cellMetaIndex.Value;
        }

        uint? valueMetaIndex = attributes.GetUintAttribute("vm");
        if (valueMetaIndex is not null)
        {
            xlCell.ValueMetaIndex = valueMetaIndex.Value;
        }

        // Move from cell start element onwards.
        reader.MoveAhead();

        bool cellHasFormula = reader.IsStartElement("f");
        XLCellFormula formula = null;
        if (cellHasFormula)
        {
            formula = this.SetCellFormula(ws, cellAddress, reader);

            // Move from end of 'f' element.
            reader.MoveAhead();
        }

        // Unified code to load value. Value can be empty and only type specified (e.g. when formula doesn't save values)
        // String type is only for formulas, while shared string/inline string/date is only for pure cell values.
        bool cellHasValue = reader.IsStartElement("v");
        if (cellHasValue)
        {
            this.SetCellValue(dataType, reader.GetText(), xlCell, cellFormat, sharedStrings);

            // Skips all nodes of the 'v' element (has no child nodes) and moves to the first element after.
            reader.Skip();
        }
        else
        {
            // A string cell must contain at least empty string.
            if (dataType.Equals(CellValues.SharedString) || dataType.Equals(CellValues.String))
            {
                xlCell.SetOnlyValue(string.Empty);
            }
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when value is missing. Formula can be null in some cases, e.g. slave cells of array formula.
        if (formula is not null && !cellHasValue)
        {
            formula.IsDirty = true;
        }

        // Inline text is dealt separately, because it is in a separate element.
        bool cellHasInlineString = reader.IsStartElement("is");
        if (cellHasInlineString)
        {
            if (dataType == CellValues.InlineString)
            {
                xlCell.ShareString = false;
                RstType inlineString = (RstType)reader.LoadCurrentElement();
                if (inlineString is not null)
                {
                    if (inlineString.Text is not null)
                    {
                        xlCell.SetOnlyValue(inlineString.Text.Text.FixNewLines());
                    }
                    else
                    {
                        this.SetCellText(xlCell, inlineString);
                    }
                }
                else
                {
                    xlCell.SetOnlyValue(string.Empty);
                }

                // Move from end 'is' element to the end of a 'c' element.
                reader.MoveAhead();
            }
            else
            {
                // Move to the first node after end of 'is' element, which should be end of cell.
                reader.Skip();
            }
        }

        if (ws.Workbook.Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
        {
            // Internally XlsxSharp stores cells as standard 1900-based style
            // so if a workbook is in 1904-format, we do that adjustment here and when saving.
            xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
        }
    }

    private XLCellFormula SetCellFormula(
        XLWorksheet ws,
        Point cellAddress,
        OpenXmlPartReader reader
    )
    {
        ReadOnlyCollection<OpenXmlAttribute> attributes = reader.Attributes;
        FormulaSlice formulaSlice = ws.Internals.CellsCollection.FormulaSlice;
        ValueSlice valueSlice = ws.Internals.CellsCollection.ValueSlice;

        // bx attribute of cell formula is not ever used, per MS-OI29500 2.1.620
        string formulaText = reader.GetText();
        CellFormulaValues formulaType = attributes.GetAttribute("t") switch
        {
            "normal" => CellFormulaValues.Normal,
            "array" => CellFormulaValues.Array,
            "dataTable" => CellFormulaValues.DataTable,
            "shared" => CellFormulaValues.Shared,
            null => CellFormulaValues.Normal,
            _ => throw new NotSupportedException("Unknown formula type."),
        };

        // Always set shareString flag to `false`, because the text result of
        // formula is stored directly in the sheet, not shared string table.
        XLCellFormula formula = null;
        if (formulaType == CellFormulaValues.Normal)
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.Set(cellAddress, formula);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (
            formulaType == CellFormulaValues.Array
            && attributes.GetRefAttribute("ref") is { } arrayArea
        ) // Child cells of an array may have array type, but not ref, that is reserved for master cell
        {
            bool aca = attributes.GetBoolAttribute("aca", false);

            // Because cells are read from top-to-bottom, from left-to-right, none of child cells have
            // a formula yet. Also, Excel doesn't allow change of array data, only through parent formula.
            formula = XLCellFormula.Array(formulaText, arrayArea, aca);
            formulaSlice.SetArray(arrayArea, formula);

            for (int col = arrayArea.FirstPoint.Column; col <= arrayArea.LastPoint.Column; ++col)
            {
                for (int row = arrayArea.FirstPoint.Row; row <= arrayArea.LastPoint.Row; ++row)
                {
                    valueSlice.SetShareString(cellAddress, false);
                }
            }
        }
        else if (
            formulaType == CellFormulaValues.Shared
            && attributes.GetUintAttribute("si") is { } sharedIndex
        )
        {
            // Shared formulas are rather limited in use and parsing, even by Excel
            // https://stackoverflow.com/questions/54654993. Therefore we accept them,
            // but don't output them. Shared formula is created, when user in Excel
            // takes a supported formula and drags it to more cells.
            if (!this._sharedFormulasR1C1.TryGetValue(sharedIndex, out string sharedR1C1Formula))
            {
                // Spec: The first formula in a group of shared formulas is saved
                // in the f element. This is considered the 'master' formula cell.
                formula = XLCellFormula.NormalA1(formulaText);
                formulaSlice.Set(cellAddress, formula);

                // The key reason why Excel hates shared formulas is likely relative addressing and the messy situation it creates
                string formulaR1C1 = FormulaConverter.ToR1C1(
                    formulaText,
                    cellAddress.Row,
                    cellAddress.Column
                );
                this._sharedFormulasR1C1.Add(sharedIndex, formulaR1C1);
            }
            else
            {
                // Spec: The formula expression for a cell that is specified to be part of a shared formula
                // (and is not the master) shall be ignored, and the master formula shall override.
                string sharedFormulaA1 = FormulaConverter.ToA1(
                    sharedR1C1Formula,
                    cellAddress.Row,
                    cellAddress.Column
                );
                formula = XLCellFormula.NormalA1(sharedFormulaA1);
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }
        else if (
            formulaType == CellFormulaValues.DataTable
            && attributes.GetRefAttribute("ref") is { } dataTableArea
        )
        {
            bool is2D = attributes.GetBoolAttribute("dt2D", false);
            bool input1Deleted = attributes.GetBoolAttribute("del1", false);
            Point input1 =
                attributes.GetCellRefAttribute("r1")
                ?? throw PartStructureException.MissingAttribute("r1");
            if (is2D)
            {
                // Input 2 is only used for 2D tables
                bool input2Deleted = attributes.GetBoolAttribute("del2", false);
                Point input2 =
                    attributes.GetCellRefAttribute("r2")
                    ?? throw PartStructureException.MissingAttribute("r2");
                formula = XLCellFormula.DataTable2D(
                    dataTableArea,
                    input1,
                    input1Deleted,
                    input2,
                    input2Deleted
                );
                formulaSlice.Set(cellAddress, formula);
            }
            else
            {
                bool isRowDataTable = attributes.GetBoolAttribute("dtr", false);
                formula = XLCellFormula.DataTable1D(
                    dataTableArea,
                    input1,
                    input1Deleted,
                    isRowDataTable
                );
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }

        // Go from start of 'f' element to the end of 'f' element.
        reader.MoveAhead();

        return formula;
    }

    private void SetCellValue(
        CellValues dataType,
        string cellValue,
        XLCell xlCell,
        XLCellFormatValue format,
        SharedStringItem[] sharedStrings
    )
    {
        if (dataType == CellValues.Number)
        {
            // XLCell is by default blank, so no need to set it.
            if (
                cellValue is not null
                && double.TryParse(
                    cellValue,
                    XlsxSharp.XLHelper.NumberStyle,
                    XlsxSharp.XLHelper.ParseCulture,
                    out double number
                )
            )
            {
                XLDataType numberDataType = format.NumberFormat.GetNumberDataType();
                XLCellValue cellNumber = numberDataType switch
                {
                    XLDataType.DateTime => XLCellValue.FromSerialDateTime(number),
                    XLDataType.TimeSpan => XLCellValue.FromSerialTimeSpan(number),
                    _ => number, // Normal number
                };
                xlCell.SetOnlyValue(cellNumber);
            }
        }
        else if (dataType == CellValues.SharedString)
        {
            if (
                cellValue is not null
                && int.TryParse(
                    cellValue,
                    XlsxSharp.XLHelper.NumberStyle,
                    XlsxSharp.XLHelper.ParseCulture,
                    out int sharedStringId
                )
                && sharedStringId >= 0
                && sharedStringId < sharedStrings.Length
            )
            {
                SharedStringItem sharedString = sharedStrings[sharedStringId];

                this.SetCellText(xlCell, sharedString);
            }
            else
            {
                xlCell.SetOnlyValue(string.Empty);
            }
        }
        else if (dataType == CellValues.String) // A plain string that is a result of a formula calculation
        {
            xlCell.SetOnlyValue(cellValue ?? string.Empty);
        }
        else if (dataType == CellValues.Boolean)
        {
            if (cellValue is not null)
            {
                bool isTrue =
                    string.Equals(cellValue, "1", StringComparison.Ordinal)
                    || string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
                xlCell.SetOnlyValue(isTrue);
            }
        }
        else if (dataType == CellValues.Error)
        {
            if (cellValue is not null && XLErrorParser.TryParseError(cellValue, out XLError error))
            {
                xlCell.SetOnlyValue(error);
            }
        }
        else if (dataType == CellValues.Date)
        {
            // Technically, cell can contain date as ISO8601 string, but not rarely used due
            // to inconsistencies between ISO and serial date time representation.
            if (cellValue is not null)
            {
                DateTime date = DateTime.ParseExact(
                    cellValue,
                    DateCellFormats,
                    XlsxSharp.XLHelper.ParseCulture,
                    DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite
                );
                xlCell.SetOnlyValue(date);
            }
        }
    }

    /// <summary>
    /// Parses the cell value for normal or rich text
    /// Input element should either be a shared string or inline string
    /// </summary>
    /// <param name="xlCell">The cell.</param>
    /// <param name="element">The element (either a shared string or inline string)</param>
    private void SetCellText(XLCell xlCell, RstType element)
    {
        // TODO Styles: Create XLImmutableRichText and assign directly instead of using the API.
        IEnumerable<Run> runs = element.Elements<Run>();
        bool hasRuns = false;
        foreach (Run run in runs)
        {
            hasRuns = true;
            RunProperties runProperties = run.RunProperties;
            string text = run.Text.InnerText.FixNewLines();

            if (runProperties == null)
            {
                xlCell.GetRichText().AddText(text, xlCell.Style.Font);
            }
            else
            {
                IXLRichString rt = xlCell.GetRichText().AddText(text);
                FontScheme fontScheme = runProperties.Elements<FontScheme>().FirstOrDefault();
                if (fontScheme != null && fontScheme.Val is not null)
                {
                    rt.SetFontScheme(fontScheme.Val.Value.ToXlsxSharp());
                }

                OpenXmlHelper.LoadFont(runProperties, rt);
            }
        }

        if (!hasRuns)
        {
            xlCell.SetOnlyValue(XStringConvert.Decode(element.Text?.InnerText) ?? string.Empty);
        }

        // Load phonetic properties
        IEnumerable<PhoneticProperties> phoneticProperties = element.Elements<PhoneticProperties>();
        PhoneticProperties pp = phoneticProperties.FirstOrDefault();
        if (pp != null)
        {
            XLPhonetics xlPhoneticPr = xlCell.GetRichText().Phonetics;

            if (pp.Alignment != null)
            {
                xlPhoneticPr.Alignment = pp.Alignment.Value.ToXlsxSharp();
            }

            if (pp.Type != null)
            {
                xlPhoneticPr.Type = pp.Type.Value.ToXlsxSharp();
            }

            if (pp.FontId?.Value is { } fontId)
            {
                XLFontFormatValue phoneticsFont = xlCell.Worksheet.Workbook.Styles.Fonts[
                    checked((int)fontId)
                ];

                xlPhoneticPr.Bold = phoneticsFont.Bold;
                xlPhoneticPr.Italic = phoneticsFont.Italic;
                xlPhoneticPr.Underline = phoneticsFont.Underline;
                xlPhoneticPr.Strikethrough = phoneticsFont.Strikethrough;
                xlPhoneticPr.VerticalAlignment = phoneticsFont.VerticalAlignment;
                xlPhoneticPr.Shadow = phoneticsFont.Shadow;
                xlPhoneticPr.FontSize = phoneticsFont.Size.Points;
                xlPhoneticPr.FontColor = phoneticsFont.Color;
                xlPhoneticPr.FontName = phoneticsFont.Name.Text;
                xlPhoneticPr.FontFamilyNumbering = phoneticsFont.Family;
                xlPhoneticPr.FontCharSet = phoneticsFont.Charset;
                xlPhoneticPr.FontScheme = phoneticsFont.Scheme;
            }
        }

        // Load phonetic runs
        IEnumerable<PhoneticRun> phoneticRuns = element.Elements<PhoneticRun>();
        foreach (PhoneticRun pr in phoneticRuns)
        {
            xlCell
                .GetRichText()
                .Phonetics.Add(
                    pr.Text.InnerText.FixNewLines(),
                    (int)pr.BaseTextStartIndex.Value,
                    (int)pr.EndingBaseIndex.Value
                );
        }
    }

    private static void LoadSheetViews(XElement sheetViews, XLWorksheet ws)
    {
        XElement sheetView = sheetViews
            ?.Elements(SpreadsheetXml.Main + "sheetView")
            .FirstOrDefault();

        if (sheetView is null)
        {
            return;
        }

        if (SpreadsheetXml.Bool(sheetView, "rightToLeft") is { } rightToLeft)
        {
            ws.RightToLeft = rightToLeft;
        }

        if (SpreadsheetXml.Bool(sheetView, "showFormulas") is { } showFormulas)
        {
            ws.ShowFormulas = showFormulas;
        }

        if (SpreadsheetXml.Bool(sheetView, "showGridLines") is { } showGridLines)
        {
            ws.ShowGridLines = showGridLines;
        }

        if (SpreadsheetXml.Bool(sheetView, "showOutlineSymbols") is { } showOutlineSymbols)
        {
            ws.ShowOutlineSymbols = showOutlineSymbols;
        }

        if (SpreadsheetXml.Bool(sheetView, "showRowColHeaders") is { } showRowColHeaders)
        {
            ws.ShowRowColHeaders = showRowColHeaders;
        }

        if (SpreadsheetXml.Bool(sheetView, "showRuler") is { } showRuler)
        {
            ws.ShowRuler = showRuler;
        }

        if (SpreadsheetXml.Bool(sheetView, "showWhiteSpace") is { } showWhiteSpace)
        {
            ws.ShowWhiteSpace = showWhiteSpace;
        }

        if (SpreadsheetXml.Bool(sheetView, "showZeros") is { } showZeros)
        {
            ws.ShowZeros = showZeros;
        }

        if (SpreadsheetXml.Bool(sheetView, "tabSelected") is { } tabSelected)
        {
            ws.TabSelected = tabSelected;
        }

        if (sheetView.Elements(SpreadsheetXml.Main + "selection").FirstOrDefault() is { } selection)
        {
            if (SpreadsheetXml.String(selection, "sqref") is { } references)
            {
                ws.Ranges(references.Replace(" ", ",")).Select();
            }

            if (SpreadsheetXml.String(selection, "activeCell") is { } activeCell)
            {
                ws.Cell(activeCell).SetActive();
            }
        }

        if (SpreadsheetXml.UInt(sheetView, "zoomScale") is { } zoomScale)
        {
            ws.SheetView.ZoomScale = (int)zoomScale;
        }

        if (SpreadsheetXml.UInt(sheetView, "zoomScaleNormal") is { } zoomScaleNormal)
        {
            ws.SheetView.ZoomScaleNormal = (int)zoomScaleNormal;
        }

        if (SpreadsheetXml.UInt(sheetView, "zoomScalePageLayoutView") is { } zoomScalePageLayout)
        {
            ws.SheetView.ZoomScalePageLayoutView = (int)zoomScalePageLayout;
        }

        if (SpreadsheetXml.UInt(sheetView, "zoomScaleSheetLayoutView") is { } zoomScaleSheetLayout)
        {
            ws.SheetView.ZoomScaleSheetLayoutView = (int)zoomScaleSheetLayout;
        }

        // A split that is not frozen is a pair of scrollbars, not a frozen pane, and the workbook
        // model has nowhere to put it.
        XElement pane = sheetView.Elements(SpreadsheetXml.Main + "pane").FirstOrDefault();
        if (SpreadsheetXml.String(pane, "state") is "frozen" or "frozenSplit")
        {
            if (SpreadsheetXml.Double(pane, "xSplit") is { } horizontalSplit)
            {
                ws.SheetView.SplitColumn = (int)horizontalSplit;
            }

            if (SpreadsheetXml.Double(pane, "ySplit") is { } verticalSplit)
            {
                ws.SheetView.SplitRow = (int)verticalSplit;
            }
        }

        string topLeftCell = SpreadsheetXml.String(sheetView, "topLeftCell");
        if (XlsxSharp.XLHelper.IsValidA1Address(topLeftCell))
        {
            ws.SheetView.TopLeftCellAddress = ws.Cell(topLeftCell).Address;
        }
    }

    private static void LoadSheetProtection(XElement sp, XLWorksheet ws)
    {
        if (sp is null)
        {
            return;
        }

        ws.Protection.IsProtected = SpreadsheetXml.Bool(sp, "sheet") ?? false;

        string algorithmName = SpreadsheetXml.String(sp, "algorithmName") ?? string.Empty;
        if (string.IsNullOrEmpty(algorithmName))
        {
            ws.Protection.PasswordHash = SpreadsheetXml.String(sp, "password") ?? string.Empty;
            ws.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (
            DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName)
        )
        {
            ws.Protection.Algorithm =
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            ws.Protection.PasswordHash = SpreadsheetXml.String(sp, "hashValue") ?? string.Empty;
            ws.Protection.SpinCount = SpreadsheetXml.UInt(sp, "spinCount") ?? 0;
            ws.Protection.Base64EncodedSalt =
                SpreadsheetXml.String(sp, "saltValue") ?? string.Empty;
        }

        // Every one of these attributes says what is *denied*, so the workbook model allows the
        // element when the attribute is off. They differ only in what an absent attribute means.
        Deny(XLSheetProtectionElements.FormatCells, "formatCells", true);
        Deny(XLSheetProtectionElements.FormatColumns, "formatColumns", true);
        Deny(XLSheetProtectionElements.FormatRows, "formatRows", true);
        Deny(XLSheetProtectionElements.InsertColumns, "insertColumns", true);
        Deny(XLSheetProtectionElements.InsertHyperlinks, "insertHyperlinks", true);
        Deny(XLSheetProtectionElements.InsertRows, "insertRows", true);
        Deny(XLSheetProtectionElements.DeleteColumns, "deleteColumns", true);
        Deny(XLSheetProtectionElements.DeleteRows, "deleteRows", true);
        Deny(XLSheetProtectionElements.AutoFilter, "autoFilter", true);
        Deny(XLSheetProtectionElements.PivotTables, "pivotTables", true);
        Deny(XLSheetProtectionElements.Sort, "sort", true);
        Deny(XLSheetProtectionElements.EditScenarios, "scenarios", true);
        Deny(XLSheetProtectionElements.EditObjects, "objects", false);
        Deny(XLSheetProtectionElements.SelectLockedCells, "selectLockedCells", false);
        Deny(XLSheetProtectionElements.SelectUnlockedCells, "selectUnlockedCells", false);

        void Deny(XLSheetProtectionElements element, string attributeName, bool deniedByDefault) =>
            ws.Protection.AllowElement(
                element,
                !(SpreadsheetXml.Bool(sp, attributeName) ?? deniedByDefault)
            );
    }

    /// <summary>
    /// Loads the conditional formatting.
    /// </summary>
    // https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.conditionalformattingrule%28v=office.15%29.aspx?f=255&MSPPError=-2147217396
    private static void LoadConditionalFormatting(
        XElement conditionalFormatting,
        XLWorksheet ws,
        LoadContext context
    )
    {
        if (conditionalFormatting is null)
        {
            return;
        }

        IReadOnlyBiDictionary<int, XLDxfValue> differentialFormats = ws.Workbook
            .Styles
            .DifferentialFormats;
        foreach (XElement fr in conditionalFormatting.Elements(SpreadsheetXml.Main + "cfRule"))
        {
            // The reference is a whitespace separated list, which is what the SDK's list value
            // parsed it as.
            IEnumerable<XLRange> ranges = SpreadsheetXml
                .String(conditionalFormatting, "sqref")
                .Split((char[])null, StringSplitOptions.RemoveEmptyEntries)
                .Select(ws.Range);
            XLConditionalFormat conditionalFormat = new(ws, XLAreaList.FromRanges(ws, ranges));

            conditionalFormat.StopIfTrue = SpreadsheetXml.Bool(fr, "stopIfTrue") ?? false;

            // TODO Styles: CF with empty format is technically legal, but seriously suss. Investigate.
            if (SpreadsheetXml.UInt(fr, "dxfId") is { } formatId)
            {
                conditionalFormat.FormatValue = differentialFormats[checked((int)formatId)];
            }

            // The conditional formatting type is compulsory. If it doesn't exist, skip the entire rule.
            if (SpreadsheetXml.String(fr, "type") is not { } type)
            {
                continue;
            }

            conditionalFormat.ConditionalFormatType = WorksheetXmlEnums.ParseConditionalFormatType(
                type
            );
            conditionalFormat.Priority = SpreadsheetXml.Int(fr, "priority") ?? int.MaxValue;

            // Although formulas are directly used only by CellIs and Expression type, other
            // format types also write them for evaluation to the workbook, e.g. rule to
            // IsBlank writes `LEN(TRIM(A2))=0` or ContainsText writes `NOT(ISERROR(SEARCH("hello",A2)))`.
            if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.CellIs)
            {
                conditionalFormat.Operator = WorksheetXmlEnums.ParseCfOperator(
                    SpreadsheetXml.String(fr, "operator")
                );

                // The XML schema allows up to three <formula> tags, but at most two are used.
                // Some producers emit empty <formula> tags that should be ignored and extra
                // non-empty formulas should also be ignored (Excel behavior).
                List<XLFormula> nonEmptyFormulas = [.. NonEmptyFormulas(fr).Select(GetFormula)];
                if (conditionalFormat.Operator is XLCFOperator.Between or XLCFOperator.NotBetween)
                {
                    List<XLFormula> formulas = [.. nonEmptyFormulas.Take(2)];
                    if (formulas.Count != 2)
                    {
                        throw PartStructureException.IncorrectElementsCount();
                    }

                    conditionalFormat.Values.Add(formulas[0]);
                    conditionalFormat.Values.Add(formulas[1]);
                }
                else
                {
                    // Other XLCFOperators expect one argument.
                    XLFormula operatorArg = nonEmptyFormulas.FirstOrDefault();
                    if (operatorArg is null)
                    {
                        throw PartStructureException.IncorrectElementsCount();
                    }

                    conditionalFormat.Values.Add(operatorArg);
                }
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Expression)
            {
                if (NonEmptyFormulas(fr).FirstOrDefault() is not { } formula)
                {
                    throw PartStructureException.IncorrectElementsCount();
                }

                conditionalFormat.Values.Add(GetFormula(formula));
            }
            else if (
                conditionalFormat.ConditionalFormatType
                is XLConditionalFormatType.ContainsText
                    or XLConditionalFormatType.NotContainsText
                    or XLConditionalFormatType.StartsWith
                    or XLConditionalFormatType.EndsWith
            )
            {
                conditionalFormat.Values.Add(
                    new XLFormula(SpreadsheetXml.String(fr, "text") ?? string.Empty)
                    {
                        IsFormula = false,
                    }
                );
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.Top10)
            {
                if (SpreadsheetXml.Bool(fr, "percent") is { } percent)
                {
                    conditionalFormat.Percent = percent;
                }

                if (SpreadsheetXml.Bool(fr, "bottom") is { } bottom)
                {
                    conditionalFormat.Bottom = bottom;
                }

                if (SpreadsheetXml.UInt(fr, "rank") is { } rank)
                {
                    conditionalFormat.Values.Add(
                        GetFormula(rank.ToString(CultureInfo.InvariantCulture))
                    );
                }
            }
            else if (conditionalFormat.ConditionalFormatType == XLConditionalFormatType.TimePeriod)
            {
                conditionalFormat.TimePeriod = SpreadsheetXml.String(fr, "timePeriod")
                    is { } timePeriod
                    ? WorksheetXmlEnums.ParseTimePeriod(timePeriod)
                    : XLTimePeriod.Yesterday;
            }

            if (fr.Element(SpreadsheetXml.Main + "colorScale") is { } colorScale)
            {
                ExtractConditionalFormatValueObjects(conditionalFormat, colorScale);
            }
            else if (fr.Element(SpreadsheetXml.Main + "dataBar") is { } dataBar)
            {
                if (SpreadsheetXml.Bool(dataBar, "showValue") is { } showValue)
                {
                    conditionalFormat.ShowBarOnly = !showValue;
                }

                // The x14 extension of a data bar carries the guid that ties this rule to the
                // richer x14 rule further down the part, written as "{...}".
                string id = fr.Descendants(SpreadsheetXml.X14 + "id").FirstOrDefault()?.Value;
                if (!string.IsNullOrWhiteSpace(id))
                {
                    conditionalFormat.Id = new Guid(id[1..^1]);
                }

                ExtractConditionalFormatValueObjects(conditionalFormat, dataBar);
            }
            else if (fr.Element(SpreadsheetXml.Main + "iconSet") is { } iconSet)
            {
                if (SpreadsheetXml.Bool(iconSet, "showValue") is { } showValue)
                {
                    conditionalFormat.ShowIconOnly = !showValue;
                }

                if (SpreadsheetXml.Bool(iconSet, "reverse") is { } reverse)
                {
                    conditionalFormat.ReverseIconOrder = reverse;
                }

                conditionalFormat.IconSetStyle = SpreadsheetXml.String(iconSet, "iconSet")
                    is { } iconSetStyle
                    ? WorksheetXmlEnums.ParseIconSetStyle(iconSetStyle)
                    : XLIconSetStyle.ThreeTrafficLights1;

                ExtractConditionalFormatValueObjects(conditionalFormat, iconSet);
            }

            bool isPivotTableFormatting =
                SpreadsheetXml.Bool(conditionalFormatting, "pivot") ?? false;
            if (isPivotTableFormatting)
            {
                context.AddPivotTableCf(ws.Name, conditionalFormat);
            }
            else
            {
                ws.ConditionalFormats.Add(conditionalFormat);
            }
        }
    }

    private static IEnumerable<string> NonEmptyFormulas(XElement cfRule) =>
        cfRule
            .Elements(SpreadsheetXml.Main + "formula")
            .Select(static f => f.Value)
            .Where(static f => !string.IsNullOrEmpty(f));

    private static XLFormula GetFormula(string value)
    {
        XLFormula formula = new();
        formula._value = value;
        formula.IsFormula = !(value[0] == '"' && value.EndsWith('"'));
        return formula;
    }

    private static void ExtractConditionalFormatValueObjects(
        XLConditionalFormat conditionalFormat,
        XElement element
    )
    {
        foreach (XElement c in element.Elements(SpreadsheetXml.Main + "cfvo"))
        {
            if (SpreadsheetXml.String(c, "type") is { } type)
            {
                conditionalFormat.ContentTypes.Add(WorksheetXmlEnums.ParseCfContentType(type));
            }

            conditionalFormat.Values.Add(
                SpreadsheetXml.String(c, "val") is { } value
                    ? new XLFormula { Value = value }
                    : null
            );

            conditionalFormat.IconSetOperators.Add(
                SpreadsheetXml.Bool(c, "gte") == false
                    ? XLCFIconSetOperator.GreaterThan
                    : XLCFIconSetOperator.EqualOrGreaterThan
            );
        }

        foreach (XElement c in element.Elements(SpreadsheetXml.Main + "color"))
        {
            conditionalFormat.Colors.Add(SpreadsheetXml.ReadColor(c));
        }
    }

    private static void LoadDataValidations(XElement dataValidations, XLWorksheet ws)
    {
        if (dataValidations is null)
        {
            return;
        }

        foreach (XElement dvs in dataValidations.Elements(SpreadsheetXml.Main + "dataValidation"))
        {
            string txt = SpreadsheetXml.String(dvs, "sqref");
            if (string.IsNullOrWhiteSpace(txt))
            {
                continue;
            }

            foreach (string rangeAddress in txt.Split(' '))
            {
                XLDataValidation dvt = ws.DataValidations.Create(Area.Parse(rangeAddress));
                LoadDataValidation(dvs, dvt);
            }
        }
    }

    /// <summary>
    /// The attributes of one <c>dataValidation</c>, which the x14 extension list repeats verbatim
    /// in its own namespace for the validations that reference another sheet.
    /// </summary>
    private static void LoadDataValidation(XElement dvs, XLDataValidation dvt)
    {
        XNamespace ns = dvs.Name.Namespace;

        if (SpreadsheetXml.Bool(dvs, "allowBlank") is { } allowBlank)
        {
            dvt.IgnoreBlanks = allowBlank;
        }

        if (SpreadsheetXml.Bool(dvs, "showDropDown") is { } showDropDown)
        {
            dvt.InCellDropdown = !showDropDown;
        }

        if (SpreadsheetXml.Bool(dvs, "showErrorMessage") is { } showErrorMessage)
        {
            dvt.ShowErrorMessage = showErrorMessage;
        }

        if (SpreadsheetXml.Bool(dvs, "showInputMessage") is { } showInputMessage)
        {
            dvt.ShowInputMessage = showInputMessage;
        }

        if (SpreadsheetXml.String(dvs, "promptTitle") is { } promptTitle)
        {
            dvt.InputTitle = promptTitle;
        }

        if (SpreadsheetXml.String(dvs, "prompt") is { } prompt)
        {
            dvt.InputMessage = prompt;
        }

        if (SpreadsheetXml.String(dvs, "errorTitle") is { } errorTitle)
        {
            dvt.ErrorTitle = errorTitle;
        }

        if (SpreadsheetXml.String(dvs, "error") is { } error)
        {
            dvt.ErrorMessage = error;
        }

        if (SpreadsheetXml.String(dvs, "errorStyle") is { } errorStyle)
        {
            dvt.ErrorStyle = WorksheetXmlEnums.ParseErrorStyle(errorStyle);
        }

        if (SpreadsheetXml.String(dvs, "type") is { } type)
        {
            dvt.AllowedValues = WorksheetXmlEnums.ParseAllowedValues(type);
        }

        if (SpreadsheetXml.String(dvs, "operator") is { } op)
        {
            dvt.Operator = WorksheetXmlEnums.ParseDataValidationOperator(op);
        }

        if (dvs.Element(ns + "formula1") is { } formula1)
        {
            dvt.MinValue = formula1.Value;
        }

        if (dvs.Element(ns + "formula2") is { } formula2)
        {
            dvt.MaxValue = formula2.Value;
        }
    }

    private static void LoadHyperlinks(
        XElement hyperlinks,
        WorksheetPart worksheetPart,
        XLWorksheet ws
    )
    {
        if (hyperlinks is null)
        {
            return;
        }

        Dictionary<string, Uri> hyperlinkDictionary =
            worksheetPart.HyperlinkRelationships?.ToDictionary(hr => hr.Id, hr => hr.Uri) ?? [];

        foreach (XElement hl in hyperlinks.Elements(SpreadsheetXml.Main + "hyperlink"))
        {
            string reference = SpreadsheetXml.String(hl, "ref");
            if (reference.Equals("#REF", StringComparison.Ordinal))
            {
                continue;
            }

            string tooltip = SpreadsheetXml.String(hl, "tooltip") ?? string.Empty;
            string relId = hl.Attribute(SpreadsheetXml.Rel + "id")?.Value;
            string location = SpreadsheetXml.String(hl, "location");
            XLRange xlRange = ws.Range(reference);
            foreach (XLCell xlCell in xlRange.Cells())
            {
                if (relId is not null)
                {
                    xlCell.SetCellHyperlink(new XLHyperlink(hyperlinkDictionary[relId], tooltip));
                }
                else if (location is not null)
                {
                    xlCell.SetCellHyperlink(new XLHyperlink(location, tooltip));
                }
                else
                {
                    xlCell.SetCellHyperlink(new XLHyperlink(reference, tooltip));
                }
            }
        }
    }

    private static void LoadPrintOptions(XElement printOptions, XLWorksheet ws)
    {
        if (printOptions is null)
        {
            return;
        }

        if (SpreadsheetXml.Bool(printOptions, "gridLines") is { } gridLines)
        {
            ws.PageSetup.ShowGridlines = gridLines;
        }

        if (SpreadsheetXml.Bool(printOptions, "horizontalCentered") is { } horizontalCentered)
        {
            ws.PageSetup.CenterHorizontally = horizontalCentered;
        }

        if (SpreadsheetXml.Bool(printOptions, "verticalCentered") is { } verticalCentered)
        {
            ws.PageSetup.CenterVertically = verticalCentered;
        }

        if (SpreadsheetXml.Bool(printOptions, "headings") is { } headings)
        {
            ws.PageSetup.ShowRowAndColumnHeadings = headings;
        }
    }

    private static void LoadPageMargins(XElement pageMargins, XLWorksheet ws)
    {
        if (pageMargins is null)
        {
            return;
        }

        IXLMargins margins = ws.PageSetup.Margins;

        if (SpreadsheetXml.Double(pageMargins, "bottom") is { } bottom)
        {
            margins.Bottom = bottom;
        }

        if (SpreadsheetXml.Double(pageMargins, "footer") is { } footer)
        {
            margins.Footer = footer;
        }

        if (SpreadsheetXml.Double(pageMargins, "header") is { } header)
        {
            margins.Header = header;
        }

        if (SpreadsheetXml.Double(pageMargins, "left") is { } left)
        {
            margins.Left = left;
        }

        if (SpreadsheetXml.Double(pageMargins, "right") is { } right)
        {
            margins.Right = right;
        }

        if (SpreadsheetXml.Double(pageMargins, "top") is { } top)
        {
            margins.Top = top;
        }
    }

    private static void LoadPageSetup(
        XElement pageSetup,
        XLWorksheet ws,
        XElement pageSetupProperties
    )
    {
        if (pageSetup is null)
        {
            return;
        }

        if (SpreadsheetXml.Int(pageSetup, "paperSize") is { } paperSize)
        {
            ws.PageSetup.PaperSize = (XLPaperSize)paperSize;
        }

        if (SpreadsheetXml.Int(pageSetup, "scale") is { } scale)
        {
            ws.PageSetup.Scale = scale;
        }

        // Both counts default to one page, so a sheet that is set to fit to page but names
        // neither fits on a single page.
        if (SpreadsheetXml.Bool(pageSetupProperties, "fitToPage") ?? false)
        {
            ws.PageSetup.PagesWide = SpreadsheetXml.Int(pageSetup, "fitToWidth") ?? 1;
            ws.PageSetup.PagesTall = SpreadsheetXml.Int(pageSetup, "fitToHeight") ?? 1;
        }

        if (SpreadsheetXml.String(pageSetup, "pageOrder") is { } pageOrder)
        {
            ws.PageSetup.PageOrder = WorksheetXmlEnums.ParsePageOrder(pageOrder);
        }

        if (SpreadsheetXml.String(pageSetup, "orientation") is { } orientation)
        {
            ws.PageSetup.PageOrientation = WorksheetXmlEnums.ParsePageOrientation(orientation);
        }

        if (SpreadsheetXml.Bool(pageSetup, "blackAndWhite") is { } blackAndWhite)
        {
            ws.PageSetup.BlackAndWhite = blackAndWhite;
        }

        if (SpreadsheetXml.Bool(pageSetup, "draft") is { } draft)
        {
            ws.PageSetup.DraftQuality = draft;
        }

        if (SpreadsheetXml.String(pageSetup, "cellComments") is { } cellComments)
        {
            ws.PageSetup.ShowComments = WorksheetXmlEnums.ParseShowComments(cellComments);
        }

        if (SpreadsheetXml.String(pageSetup, "errors") is { } errors)
        {
            ws.PageSetup.PrintErrorValue = WorksheetXmlEnums.ParsePrintError(errors);
        }

        if (SpreadsheetXml.UInt(pageSetup, "horizontalDpi") is { } horizontalDpi)
        {
            ws.PageSetup.HorizontalDpi = (int)horizontalDpi;
        }

        if (SpreadsheetXml.UInt(pageSetup, "verticalDpi") is { } verticalDpi)
        {
            ws.PageSetup.VerticalDpi = (int)verticalDpi;
        }

        // The page number is unsigned in the schema, and a file that writes a negative one - as
        // one of the test workbooks does - leaves the sheet numbering from its own default.
        if (SpreadsheetXml.UInt(pageSetup, "firstPageNumber") is { } firstPageNumber)
        {
            ws.PageSetup.FirstPageNumber = (int)firstPageNumber;
        }
    }

    private static void LoadHeaderFooter(XElement headerFooter, XLWorksheet ws)
    {
        if (headerFooter is null)
        {
            return;
        }

        if (SpreadsheetXml.Bool(headerFooter, "alignWithMargins") is { } alignWithMargins)
        {
            ws.PageSetup.AlignHFWithMargins = alignWithMargins;
        }

        if (SpreadsheetXml.Bool(headerFooter, "scaleWithDoc") is { } scaleWithDoc)
        {
            ws.PageSetup.ScaleHFWithDocument = scaleWithDoc;
        }

        if (SpreadsheetXml.Bool(headerFooter, "differentFirst") is { } differentFirst)
        {
            ws.PageSetup.DifferentFirstPageOnHF = differentFirst;
        }

        if (SpreadsheetXml.Bool(headerFooter, "differentOddEven") is { } differentOddEven)
        {
            ws.PageSetup.DifferentOddEvenPagesOnHF = differentOddEven;
        }

        SetHeaderFooterText(
            headerFooter,
            "evenFooter",
            ws.PageSetup.Footer,
            XLHFOccurrence.EvenPages
        );
        SetHeaderFooterText(
            headerFooter,
            "oddFooter",
            ws.PageSetup.Footer,
            XLHFOccurrence.OddPages
        );
        SetHeaderFooterText(
            headerFooter,
            "firstFooter",
            ws.PageSetup.Footer,
            XLHFOccurrence.FirstPage
        );
        SetHeaderFooterText(
            headerFooter,
            "evenHeader",
            ws.PageSetup.Header,
            XLHFOccurrence.EvenPages
        );
        SetHeaderFooterText(
            headerFooter,
            "oddHeader",
            ws.PageSetup.Header,
            XLHFOccurrence.OddPages
        );
        SetHeaderFooterText(
            headerFooter,
            "firstHeader",
            ws.PageSetup.Header,
            XLHFOccurrence.FirstPage
        );

        ((XLHeaderFooter)ws.PageSetup.Header).SetAsInitial();
        ((XLHeaderFooter)ws.PageSetup.Footer).SetAsInitial();
    }

    private static void SetHeaderFooterText(
        XElement headerFooter,
        string name,
        IXLHeaderFooter target,
        XLHFOccurrence occurrence
    )
    {
        if (headerFooter.Element(SpreadsheetXml.Main + name) is { } text)
        {
            ((XLHeaderFooter)target).SetInnerText(occurrence, text.Value);
        }
    }

    private static void LoadRowBreaks(XElement rowBreaks, XLWorksheet ws) =>
        LoadBreaks(rowBreaks, ws.PageSetup.RowBreaks);

    private static void LoadColumnBreaks(XElement columnBreaks, XLWorksheet ws) =>
        LoadBreaks(columnBreaks, ws.PageSetup.ColumnBreaks);

    private static void LoadBreaks(XElement breaks, List<int> target)
    {
        if (breaks is null)
        {
            return;
        }

        foreach (XElement brk in breaks.Elements(SpreadsheetXml.Main + "brk"))
        {
            if (SpreadsheetXml.Int(brk, "id") is { } id)
            {
                target.Add(id);
            }
        }
    }

    private static void LoadExtensions(WorksheetExtensionList extensions, XLWorksheet ws)
    {
        if (extensions == null)
        {
            return;
        }

        foreach (
            X14.DataValidation dvs in extensions
                .Descendants<X14.DataValidations>()
                .SelectMany(dataValidations => dataValidations.Descendants<X14.DataValidation>())
        )
        {
            string txt = dvs.ReferenceSequence.InnerText;
            if (string.IsNullOrWhiteSpace(txt))
            {
                continue;
            }

            foreach (string rangeAddress in txt.Split(' '))
            {
                XLDataValidation dvt = ws.DataValidations.Create(Area.Parse(rangeAddress));
                if (dvs.AllowBlank != null)
                {
                    dvt.IgnoreBlanks = dvs.AllowBlank;
                }

                if (dvs.ShowDropDown != null)
                {
                    dvt.InCellDropdown = !dvs.ShowDropDown.Value;
                }

                if (dvs.ShowErrorMessage != null)
                {
                    dvt.ShowErrorMessage = dvs.ShowErrorMessage;
                }

                if (dvs.ShowInputMessage != null)
                {
                    dvt.ShowInputMessage = dvs.ShowInputMessage;
                }

                if (dvs.PromptTitle != null)
                {
                    dvt.InputTitle = dvs.PromptTitle;
                }

                if (dvs.Prompt != null)
                {
                    dvt.InputMessage = dvs.Prompt;
                }

                if (dvs.ErrorTitle != null)
                {
                    dvt.ErrorTitle = dvs.ErrorTitle;
                }

                if (dvs.Error != null)
                {
                    dvt.ErrorMessage = dvs.Error;
                }

                if (dvs.ErrorStyle != null)
                {
                    dvt.ErrorStyle = dvs.ErrorStyle.Value.ToXlsxSharp();
                }

                if (dvs.Type != null)
                {
                    dvt.AllowedValues = dvs.Type.Value.ToXlsxSharp();
                }

                if (dvs.Operator != null)
                {
                    dvt.Operator = dvs.Operator.Value.ToXlsxSharp();
                }

                if (dvs.DataValidationForumla1 != null)
                {
                    dvt.MinValue = dvs.DataValidationForumla1.InnerText;
                }

                if (dvs.DataValidationForumla2 != null)
                {
                    dvt.MaxValue = dvs.DataValidationForumla2.InnerText;
                }
            }
        }

        foreach (
            X14.ConditionalFormattingRule conditionalFormattingRule in extensions
                .Descendants<X14.ConditionalFormattingRule>()
                .Where(cf =>
                    cf.Type != null
                    && cf.Type.HasValue
                    && cf.Type.Value == ConditionalFormatValues.DataBar
                )
        )
        {
            XLConditionalFormat xlConditionalFormat = ws
                .ConditionalFormats.Cast<XLConditionalFormat>()
                .SingleOrDefault(cf => cf.Id.WrapInBraces() == conditionalFormattingRule.Id);
            if (xlConditionalFormat != null)
            {
                X14.NegativeFillColor negativeFillColor = conditionalFormattingRule
                    .Descendants<X14.NegativeFillColor>()
                    .SingleOrDefault();
                xlConditionalFormat.Colors.Add(negativeFillColor.ToXlsxSharpColor());
            }
        }

        foreach (
            X14.SparklineGroup slg in extensions
                .Descendants<X14.SparklineGroups>()
                .SelectMany(sparklineGroups => sparklineGroups.Descendants<X14.SparklineGroup>())
        )
        {
            XLSparklineGroup xlSparklineGroup = ws.SparklineGroupsInternal.Add();

            if (slg.Formula != null)
            {
                xlSparklineGroup.DateRange = ws.Workbook.Range(slg.Formula.Text);
            }

            IXLSparklineStyle xlSparklineStyle = xlSparklineGroup.Style;
            if (slg.FirstMarkerColor != null)
            {
                xlSparklineStyle.FirstMarkerColor = slg.FirstMarkerColor.ToXlsxSharpColor();
            }

            if (slg.LastMarkerColor != null)
            {
                xlSparklineStyle.LastMarkerColor = slg.LastMarkerColor.ToXlsxSharpColor();
            }

            if (slg.HighMarkerColor != null)
            {
                xlSparklineStyle.HighMarkerColor = slg.HighMarkerColor.ToXlsxSharpColor();
            }

            if (slg.LowMarkerColor != null)
            {
                xlSparklineStyle.LowMarkerColor = slg.LowMarkerColor.ToXlsxSharpColor();
            }

            if (slg.SeriesColor != null)
            {
                xlSparklineStyle.SeriesColor = slg.SeriesColor.ToXlsxSharpColor();
            }

            if (slg.NegativeColor != null)
            {
                xlSparklineStyle.NegativeColor = slg.NegativeColor.ToXlsxSharpColor();
            }

            if (slg.MarkersColor != null)
            {
                xlSparklineStyle.MarkersColor = slg.MarkersColor.ToXlsxSharpColor();
            }

            xlSparklineGroup.Style = xlSparklineStyle;

            if (slg.DisplayHidden != null)
            {
                xlSparklineGroup.DisplayHidden = slg.DisplayHidden;
            }

            if (slg.LineWeight != null)
            {
                xlSparklineGroup.LineWeight = slg.LineWeight;
            }

            if (slg.Type != null)
            {
                xlSparklineGroup.Type = slg.Type.Value.ToXlsxSharp();
            }

            if (slg.DisplayEmptyCellsAs != null)
            {
                xlSparklineGroup.DisplayEmptyCellsAs = slg.DisplayEmptyCellsAs.Value.ToXlsxSharp();
            }

            xlSparklineGroup.ShowMarkers = XLSparklineMarkers.None;
            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Markers, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.Markers;
            }

            if (OpenXmlHelper.GetBooleanValueAsBool(slg.High, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.HighPoint;
            }

            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Low, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LowPoint;
            }

            if (OpenXmlHelper.GetBooleanValueAsBool(slg.First, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.FirstPoint;
            }

            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Last, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.LastPoint;
            }

            if (OpenXmlHelper.GetBooleanValueAsBool(slg.Negative, false))
            {
                xlSparklineGroup.ShowMarkers |= XLSparklineMarkers.NegativePoints;
            }

            if (slg.AxisColor != null)
            {
                xlSparklineGroup.HorizontalAxis.Color = XLColor.FromHtml(slg.AxisColor.Rgb.Value);
            }

            if (slg.DisplayXAxis != null)
            {
                xlSparklineGroup.HorizontalAxis.IsVisible = slg.DisplayXAxis;
            }

            if (slg.RightToLeft != null)
            {
                xlSparklineGroup.HorizontalAxis.RightToLeft = slg.RightToLeft;
            }

            if (slg.ManualMax != null)
            {
                xlSparklineGroup.VerticalAxis.ManualMax = slg.ManualMax;
            }

            if (slg.ManualMin != null)
            {
                xlSparklineGroup.VerticalAxis.ManualMin = slg.ManualMin;
            }

            if (slg.MinAxisType != null)
            {
                xlSparklineGroup.VerticalAxis.MinAxisType = slg.MinAxisType.Value.ToXlsxSharp();
            }

            if (slg.MaxAxisType != null)
            {
                xlSparklineGroup.VerticalAxis.MaxAxisType = slg.MaxAxisType.Value.ToXlsxSharp();
            }

            foreach (
                X14.Sparkline sparkline in slg.Descendants<X14.Sparklines>()
                    .SelectMany(sparklines => sparklines.Descendants<X14.Sparkline>())
            )
            {
                // The sqlref must contain exactly one ref [MS-XLSX]. Excel ignores everything after the first one.
                string refText = (sparkline.ReferenceSequence?.Text ?? string.Empty)
                    .Trim()
                    .Split(' ')[0];
                Point location = Point.Parse(refText);

                // Technically, there could be more than one sparkline per cell, so use Set instead of Add.
                xlSparklineGroup.SetSparkline(location, sparkline.Formula?.Text);
            }
        }
    }

    /// <summary>
    /// Hands a loader the element as XML while the dispatch loop above still runs on the SDK
    /// reader. It goes away with the loop.
    /// </summary>
    private static XElement AsXElement(OpenXmlElement element) =>
        element is null ? null : XElement.Parse(element.OuterXml);

    private static void ApplyStyle(
        IXLFormatContainer container,
        int styleIndex,
        XLWorkbookStyles styles
    ) => container.FormatValue = styles.CellFormats[styleIndex];

    private static void ApplyStyle(XLColumns columns, int styleIndex, XLWorkbookStyles styles)
    {
        // When loading columns we must propagate style to each column but not deeper. In other cases we do not propagate at all.
        foreach (XLColumn col in columns)
        {
            ApplyStyle(col, styleIndex, styles);
        }
    }
}

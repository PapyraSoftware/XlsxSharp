using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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

    private static readonly XmlReaderSettings ReaderSettings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        CloseInput = false,
    };

    private readonly Dictionary<uint, string> _sharedFormulasR1C1 = new();

    /// <summary>
    /// The <c>t</c> attribute of a cell, which says how to read its value.
    /// </summary>
    private enum CellType
    {
        Boolean,
        Number,
        Error,
        SharedString,
        String,
        InlineString,
        Date,
    }

    /// <summary>
    /// The <c>t</c> attribute of a cell formula.
    /// </summary>
    private enum FormulaType
    {
        Normal,
        Array,
        DataTable,
        Shared,
    }

    /// <summary>
    /// Row number of last read <c>row</c> element.
    /// </summary>
    private int _lastRow;
    private int _lastColumnNumber;

    internal void LoadWorksheet(
        XLWorksheet ws,
        WorksheetPart worksheetPart,
        XElement[] sharedStrings,
        LoadContext context
    )
    {
        XElement pageSetupProperties = null;

        this._lastRow = 0;

        using Stream stream = worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
        using XmlReader reader = XmlReader.Create(stream, ReaderSettings);

        reader.MoveToContent();
        if (reader.NodeType != XmlNodeType.Element || reader.IsEmptyElement)
        {
            return;
        }

        reader.ReadStartElement();
        while (reader.NodeType == XmlNodeType.Element)
        {
            if (reader.NamespaceURI != OpenXmlConst.Main2006SsNs)
            {
                reader.Skip();
                continue;
            }

            switch (reader.LocalName)
            {
                case "sheetPr":
                    LoadSheetProperties(ReadElement(reader), ws, out pageSetupProperties);
                    break;
                case "sheetFormatPr":
                    LoadSheetFormatProperties(ReadElement(reader), ws);
                    break;
                case "sheetViews":
                    LoadSheetViews(ReadElement(reader), ws);
                    break;
                case "cols":
                    LoadColumns(ws, ReadElement(reader));
                    break;
                case "sheetData":
                    this.LoadSheetData(ws, sharedStrings, reader);
                    break;
                case "mergeCells":
                    LoadMergeCells(ReadElement(reader), ws);
                    break;
                case "autoFilter":
                    AutoFilterReader.LoadAutoFilter(ReadElement(reader), ws);
                    break;
                case "sheetProtection":
                    LoadSheetProtection(ReadElement(reader), ws);
                    break;
                case "dataValidations":
                    LoadDataValidations(ReadElement(reader), ws);
                    break;
                case "conditionalFormatting":
                    LoadConditionalFormatting(ReadElement(reader), ws, context);
                    break;
                case "hyperlinks":
                    LoadHyperlinks(ReadElement(reader), worksheetPart, ws);
                    break;
                case "printOptions":
                    LoadPrintOptions(ReadElement(reader), ws);
                    break;
                case "pageMargins":
                    LoadPageMargins(ReadElement(reader), ws);
                    break;
                case "pageSetup":
                    LoadPageSetup(ReadElement(reader), ws, pageSetupProperties);
                    break;
                case "headerFooter":
                    LoadHeaderFooter(ReadElement(reader), ws);
                    break;
                case "rowBreaks":
                    LoadRowBreaks(ReadElement(reader), ws);
                    break;
                case "colBreaks":
                    LoadColumnBreaks(ReadElement(reader), ws);
                    break;
                case "extLst":
                    LoadExtensions(ReadElement(reader), ws);
                    break;
                case "legacyDrawing":
                    ws.LegacyDrawingId = reader.GetAttribute("id", OpenXmlConst.RelationshipsNs);
                    reader.Skip();
                    break;
                default:
                    // Everything else is passed over, custom sheet views among them: they carry
                    // an auto filter and more of their own, none of which the model keeps.
                    reader.Skip();
                    break;
            }
        }
    }

    /// <summary>
    /// Materialise the element the reader is on, and leave the reader on the node after it.
    /// </summary>
    private static XElement ReadElement(XmlReader reader) => (XElement)XNode.ReadFrom(reader);

    /// <summary>
    /// The cells are the bulk of a worksheet, so they are read straight off the stream rather
    /// than materialised the way the elements around them are.
    /// </summary>
    private void LoadSheetData(XLWorksheet ws, XElement[] sharedStrings, XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Skip();
            return;
        }

        reader.ReadStartElement();
        while (reader.NodeType == XmlNodeType.Element)
        {
            if (reader.LocalName == "row" && reader.NamespaceURI == OpenXmlConst.Main2006SsNs)
            {
                this.LoadRow(ws, sharedStrings, reader);
            }
            else
            {
                reader.Skip();
            }
        }

        reader.ReadEndElement();
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

    private void LoadRow(XLWorksheet ws, XElement[] sharedStrings, XmlReader reader)
    {
        // Row number is an optional attribute. If not specified, it should be a next row from the last read row.
        int rowIndex = Int(reader, "r") ?? ++this._lastRow;
        this._lastRow = rowIndex;

        XLRow xlRow = ws.Row(rowIndex, false);

        double? height = Double(reader, "ht");
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

        double? dyDescent = Double(reader, "dyDescent", OpenXmlConst.X14Ac2009SsNs);
        if (dyDescent is not null)
        {
            xlRow.DyDescent = dyDescent.Value;
        }

        if (Bool(reader, "hidden") ?? false)
        {
            xlRow.Hide();
        }

        if (Bool(reader, "collapsed") ?? false)
        {
            xlRow.Collapsed = true;
        }

        int? outlineLevel = Int(reader, "outlineLevel");
        if (outlineLevel is not null && outlineLevel.Value > 0)
        {
            xlRow.OutlineLevel = outlineLevel.Value;
        }

        if (Bool(reader, "ph") ?? false)
        {
            xlRow.ShowPhonetic = true;
        }

        if ((Bool(reader, "customFormat") ?? false) && Int(reader, "s") is { } styleIndex)
        {
            ApplyStyle(xlRow, styleIndex, ws.Workbook.Styles);
        }

        this._lastColumnNumber = 0;

        if (reader.IsEmptyElement)
        {
            reader.Skip();
            return;
        }

        reader.ReadStartElement();
        while (reader.NodeType == XmlNodeType.Element)
        {
            if (reader.LocalName == "c" && reader.NamespaceURI == OpenXmlConst.Main2006SsNs)
            {
                this.LoadCell(sharedStrings, ws, reader, rowIndex);
            }
            else
            {
                // In theory, row can also contain extList, just skip them.
                reader.Skip();
            }
        }

        reader.ReadEndElement();
    }

    private void LoadCell(XElement[] sharedStrings, XLWorksheet ws, XmlReader reader, int rowIndex)
    {
        Point cellAddress = CellRef(reader, "r") ?? new Point(rowIndex, this._lastColumnNumber + 1);
        this._lastColumnNumber = cellAddress.Column;

        CellType dataType = reader.GetAttribute("t") switch
        {
            "b" => CellType.Boolean,
            "n" => CellType.Number,
            "e" => CellType.Error,
            "s" => CellType.SharedString,
            "str" => CellType.String,
            "inlineStr" => CellType.InlineString,
            "d" => CellType.Date,
            null => CellType.Number,
            _ => throw new FormatException($"Unknown cell type."),
        };

        XLCell xlCell = ws.Cell(cellAddress.Row, cellAddress.Column);

        int xfId = Int(reader, "s") ?? 0;
        XLCellFormatValue cellFormat = ws.Workbook.Styles.CellFormats[xfId];
        xlCell.FormatValue = cellFormat;

        if (Bool(reader, "ph") ?? false)
        {
            xlCell.ShowPhonetic = true;
        }

        if (UInt(reader, "cm") is { } cellMetaIndex)
        {
            xlCell.CellMetaIndex = cellMetaIndex;
        }

        if (UInt(reader, "vm") is { } valueMetaIndex)
        {
            xlCell.ValueMetaIndex = valueMetaIndex;
        }

        XLCellFormula formula = null;
        string cellValue = null;
        bool cellHasValue = false;
        XElement inlineString = null;

        if (reader.IsEmptyElement)
        {
            reader.Skip();
        }
        else
        {
            reader.ReadStartElement();
            while (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.NamespaceURI != OpenXmlConst.Main2006SsNs)
                {
                    reader.Skip();
                    continue;
                }

                switch (reader.LocalName)
                {
                    case "f":
                        formula = this.SetCellFormula(ws, cellAddress, reader);
                        break;
                    case "v":
                        cellHasValue = true;
                        cellValue = reader.ReadElementContentAsString();
                        break;
                    case "is":
                        // Inline text is dealt separately, because it is in a separate element.
                        inlineString = ReadElement(reader);
                        break;
                    default:
                        reader.Skip();
                        break;
                }
            }

            reader.ReadEndElement();
        }

        // Unified code to load value. Value can be empty and only type specified (e.g. when formula doesn't save values)
        // String type is only for formulas, while shared string/inline string/date is only for pure cell values.
        if (cellHasValue)
        {
            this.SetCellValue(dataType, cellValue, xlCell, cellFormat, sharedStrings);
        }
        else if (dataType is CellType.SharedString or CellType.String)
        {
            // A string cell must contain at least empty string.
            xlCell.SetOnlyValue(string.Empty);
        }

        // If the cell doesn't contain value, we should invalidate it, otherwise rely on the stored value.
        // The value is likely more reliable. It should be set when cellFormula.CalculateCell is set or
        // when value is missing. Formula can be null in some cases, e.g. slave cells of array formula.
        if (formula is not null && !cellHasValue)
        {
            formula.IsDirty = true;
        }

        if (inlineString is not null && dataType == CellType.InlineString)
        {
            xlCell.ShareString = false;
            if (inlineString.Element(SpreadsheetXml.Main + "t") is { } text)
            {
                xlCell.SetOnlyValue(text.Value.FixNewLines());
            }
            else
            {
                this.SetCellText(xlCell, inlineString);
            }
        }

        if (ws.Workbook.Use1904DateSystem && xlCell.DataType == XLDataType.DateTime)
        {
            // Internally XlsxSharp stores cells as standard 1900-based style
            // so if a workbook is in 1904-format, we do that adjustment here and when saving.
            xlCell.SetOnlyValue(xlCell.GetDateTime().AddDays(1462));
        }
    }

    private XLCellFormula SetCellFormula(XLWorksheet ws, Point cellAddress, XmlReader reader)
    {
        FormulaSlice formulaSlice = ws.Internals.CellsCollection.FormulaSlice;
        ValueSlice valueSlice = ws.Internals.CellsCollection.ValueSlice;

        // bx attribute of cell formula is not ever used, per MS-OI29500 2.1.620
        FormulaType formulaType = reader.GetAttribute("t") switch
        {
            "normal" => FormulaType.Normal,
            "array" => FormulaType.Array,
            "dataTable" => FormulaType.DataTable,
            "shared" => FormulaType.Shared,
            null => FormulaType.Normal,
            _ => throw new NotSupportedException("Unknown formula type."),
        };

        // The attributes have to be read before the text, which moves the reader past the element.
        Area? arrayOrTableArea = Ref(reader, "ref");
        bool aca = Bool(reader, "aca") ?? false;
        uint? sharedIndex = UInt(reader, "si");
        bool is2D = Bool(reader, "dt2D") ?? false;
        bool input1Deleted = Bool(reader, "del1") ?? false;
        bool input2Deleted = Bool(reader, "del2") ?? false;
        bool isRowDataTable = Bool(reader, "dtr") ?? false;
        Point? input1 = CellRef(reader, "r1");
        Point? input2 = CellRef(reader, "r2");
        string formulaText = reader.ReadElementContentAsString();

        // Always set shareString flag to `false`, because the text result of
        // formula is stored directly in the sheet, not shared string table.
        XLCellFormula formula = null;
        if (formulaType == FormulaType.Normal)
        {
            formula = XLCellFormula.NormalA1(formulaText);
            formulaSlice.Set(cellAddress, formula);
            valueSlice.SetShareString(cellAddress, false);
        }
        else if (formulaType == FormulaType.Array && arrayOrTableArea is { } arrayArea) // Child cells of an array may have array type, but not ref, that is reserved for master cell
        {
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
        else if (formulaType == FormulaType.Shared && sharedIndex is { } sharedFormulaIndex)
        {
            // Shared formulas are rather limited in use and parsing, even by Excel
            // https://stackoverflow.com/questions/54654993. Therefore we accept them,
            // but don't output them. Shared formula is created, when user in Excel
            // takes a supported formula and drags it to more cells.
            if (
                !this._sharedFormulasR1C1.TryGetValue(
                    sharedFormulaIndex,
                    out string sharedR1C1Formula
                )
            )
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
                this._sharedFormulasR1C1.Add(sharedFormulaIndex, formulaR1C1);
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
        else if (formulaType == FormulaType.DataTable && arrayOrTableArea is { } dataTableArea)
        {
            Point firstInput = input1 ?? throw PartStructureException.MissingAttribute("r1");
            if (is2D)
            {
                // Input 2 is only used for 2D tables
                Point secondInput = input2 ?? throw PartStructureException.MissingAttribute("r2");
                formula = XLCellFormula.DataTable2D(
                    dataTableArea,
                    firstInput,
                    input1Deleted,
                    secondInput,
                    input2Deleted
                );
                formulaSlice.Set(cellAddress, formula);
            }
            else
            {
                formula = XLCellFormula.DataTable1D(
                    dataTableArea,
                    firstInput,
                    input1Deleted,
                    isRowDataTable
                );
                formulaSlice.Set(cellAddress, formula);
            }

            valueSlice.SetShareString(cellAddress, false);
        }

        return formula;
    }

    private void SetCellValue(
        CellType dataType,
        string cellValue,
        XLCell xlCell,
        XLCellFormatValue format,
        XElement[] sharedStrings
    )
    {
        if (dataType == CellType.Number)
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
        else if (dataType == CellType.SharedString)
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
                XElement sharedString = sharedStrings[sharedStringId];

                this.SetCellText(xlCell, sharedString);
            }
            else
            {
                xlCell.SetOnlyValue(string.Empty);
            }
        }
        else if (dataType == CellType.String) // A plain string that is a result of a formula calculation
        {
            xlCell.SetOnlyValue(cellValue ?? string.Empty);
        }
        else if (dataType == CellType.Boolean)
        {
            if (cellValue is not null)
            {
                bool isTrue =
                    string.Equals(cellValue, "1", StringComparison.Ordinal)
                    || string.Equals(cellValue, "TRUE", StringComparison.OrdinalIgnoreCase);
                xlCell.SetOnlyValue(isTrue);
            }
        }
        else if (dataType == CellType.Error)
        {
            if (cellValue is not null && XLErrorParser.TryParseError(cellValue, out XLError error))
            {
                xlCell.SetOnlyValue(error);
            }
        }
        else if (dataType == CellType.Date)
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
    private void SetCellText(XLCell xlCell, XElement element)
    {
        // TODO Styles: Create XLImmutableRichText and assign directly instead of using the API.
        bool hasRuns = false;
        foreach (XElement run in element.Elements(SpreadsheetXml.Main + "r"))
        {
            hasRuns = true;
            XElement runProperties = run.Element(SpreadsheetXml.Main + "rPr");
            string text = run.Element(SpreadsheetXml.Main + "t").Value.FixNewLines();

            if (runProperties is null)
            {
                xlCell.GetRichText().AddText(text, xlCell.Style.Font);
            }
            else
            {
                IXLRichString rt = xlCell.GetRichText().AddText(text);
                StyleXml.LoadFont(runProperties, rt);
            }
        }

        if (!hasRuns)
        {
            xlCell.SetOnlyValue(
                XStringConvert.Decode(element.Element(SpreadsheetXml.Main + "t")?.Value)
                    ?? string.Empty
            );
        }

        LoadPhonetics(xlCell, element);
    }

    private static void LoadPhonetics(XLCell xlCell, XElement element)
    {
        if (element.Element(SpreadsheetXml.Main + "phoneticPr") is { } pp)
        {
            XLPhonetics xlPhoneticPr = xlCell.GetRichText().Phonetics;

            if (SpreadsheetXml.String(pp, "alignment") is { } alignment)
            {
                xlPhoneticPr.Alignment = WorksheetXmlEnums.ParsePhoneticAlignment(alignment);
            }

            if (SpreadsheetXml.String(pp, "type") is { } type)
            {
                xlPhoneticPr.Type = WorksheetXmlEnums.ParsePhoneticType(type);
            }

            if (SpreadsheetXml.UInt(pp, "fontId") is { } fontId)
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

        foreach (XElement pr in element.Elements(SpreadsheetXml.Main + "rPh"))
        {
            xlCell
                .GetRichText()
                .Phonetics.Add(
                    pr.Element(SpreadsheetXml.Main + "t").Value.FixNewLines(),
                    checked((int)SpreadsheetXml.UInt(pr, "sb")),
                    checked((int)SpreadsheetXml.UInt(pr, "eb"))
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

    private static void LoadExtensions(XElement extensions, XLWorksheet ws)
    {
        if (extensions is null)
        {
            return;
        }

        LoadExtensionDataValidations(extensions, ws);
        LoadExtensionDataBarColors(extensions, ws);
        LoadSparklineGroups(extensions, ws);
    }

    /// <summary>
    /// Validations whose list points at another sheet cannot be written in the 2006 schema, so
    /// they are repeated here - same attributes, own namespace, and the reference and the
    /// formulas moved out of the attributes into xm elements.
    /// </summary>
    private static void LoadExtensionDataValidations(XElement extensions, XLWorksheet ws)
    {
        foreach (
            XElement dvs in extensions
                .Descendants(SpreadsheetXml.X14 + "dataValidations")
                .SelectMany(dataValidations =>
                    dataValidations.Descendants(SpreadsheetXml.X14 + "dataValidation")
                )
        )
        {
            string txt = dvs.Element(SpreadsheetXml.Xm + "sqref")?.Value;
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
    /// A data bar's colour for negative values has no place in the 2006 schema, so it is written
    /// here on a rule that carries the guid of the one it extends.
    /// </summary>
    private static void LoadExtensionDataBarColors(XElement extensions, XLWorksheet ws)
    {
        foreach (
            XElement rule in extensions
                .Descendants(SpreadsheetXml.X14 + "cfRule")
                .Where(cf => SpreadsheetXml.String(cf, "type") == "dataBar")
        )
        {
            string id = SpreadsheetXml.String(rule, "id");
            XLConditionalFormat xlConditionalFormat = ws
                .ConditionalFormats.Cast<XLConditionalFormat>()
                .SingleOrDefault(cf => cf.Id.WrapInBraces() == id);
            if (xlConditionalFormat is not null)
            {
                XElement negativeFillColor = rule.Descendants(
                        SpreadsheetXml.X14 + "negativeFillColor"
                    )
                    .SingleOrDefault();
                xlConditionalFormat.Colors.Add(SpreadsheetXml.ReadColor(negativeFillColor));
            }
        }
    }

    private static void LoadSparklineGroups(XElement extensions, XLWorksheet ws)
    {
        foreach (
            XElement slg in extensions
                .Descendants(SpreadsheetXml.X14 + "sparklineGroups")
                .SelectMany(sparklineGroups =>
                    sparklineGroups.Descendants(SpreadsheetXml.X14 + "sparklineGroup")
                )
        )
        {
            XLSparklineGroup xlSparklineGroup = ws.SparklineGroupsInternal.Add();

            if (slg.Element(SpreadsheetXml.Xm + "f") is { } dateRange)
            {
                xlSparklineGroup.DateRange = ws.Workbook.Range(dateRange.Value);
            }

            xlSparklineGroup.Style = LoadSparklineStyle(slg, xlSparklineGroup.Style);

            if (SpreadsheetXml.Bool(slg, "displayHidden") is { } displayHidden)
            {
                xlSparklineGroup.DisplayHidden = displayHidden;
            }

            if (SpreadsheetXml.Double(slg, "lineWeight") is { } lineWeight)
            {
                xlSparklineGroup.LineWeight = lineWeight;
            }

            if (SpreadsheetXml.String(slg, "type") is { } type)
            {
                xlSparklineGroup.Type = WorksheetXmlEnums.ParseSparklineType(type);
            }

            if (SpreadsheetXml.String(slg, "displayEmptyCellsAs") is { } displayEmptyCellsAs)
            {
                xlSparklineGroup.DisplayEmptyCellsAs = WorksheetXmlEnums.ParseDisplayBlanksAs(
                    displayEmptyCellsAs
                );
            }

            xlSparklineGroup.ShowMarkers = LoadSparklineMarkers(slg);

            if (slg.Element(SpreadsheetXml.X14 + "colorAxis") is { } axisColor)
            {
                xlSparklineGroup.HorizontalAxis.Color = XLColor.FromHtml(
                    SpreadsheetXml.String(axisColor, "rgb")
                );
            }

            if (SpreadsheetXml.Bool(slg, "displayXAxis") is { } displayXAxis)
            {
                xlSparklineGroup.HorizontalAxis.IsVisible = displayXAxis;
            }

            if (SpreadsheetXml.Bool(slg, "rightToLeft") is { } rightToLeft)
            {
                xlSparklineGroup.HorizontalAxis.RightToLeft = rightToLeft;
            }

            if (SpreadsheetXml.Double(slg, "manualMax") is { } manualMax)
            {
                xlSparklineGroup.VerticalAxis.ManualMax = manualMax;
            }

            if (SpreadsheetXml.Double(slg, "manualMin") is { } manualMin)
            {
                xlSparklineGroup.VerticalAxis.ManualMin = manualMin;
            }

            if (SpreadsheetXml.String(slg, "minAxisType") is { } minAxisType)
            {
                xlSparklineGroup.VerticalAxis.MinAxisType =
                    WorksheetXmlEnums.ParseSparklineAxisMinMax(minAxisType);
            }

            if (SpreadsheetXml.String(slg, "maxAxisType") is { } maxAxisType)
            {
                xlSparklineGroup.VerticalAxis.MaxAxisType =
                    WorksheetXmlEnums.ParseSparklineAxisMinMax(maxAxisType);
            }

            LoadSparklines(slg, xlSparklineGroup);
        }
    }

    private static IXLSparklineStyle LoadSparklineStyle(
        XElement slg,
        IXLSparklineStyle xlSparklineStyle
    )
    {
        SetColor("colorFirst", c => xlSparklineStyle.FirstMarkerColor = c);
        SetColor("colorLast", c => xlSparklineStyle.LastMarkerColor = c);
        SetColor("colorHigh", c => xlSparklineStyle.HighMarkerColor = c);
        SetColor("colorLow", c => xlSparklineStyle.LowMarkerColor = c);
        SetColor("colorSeries", c => xlSparklineStyle.SeriesColor = c);
        SetColor("colorNegative", c => xlSparklineStyle.NegativeColor = c);
        SetColor("colorMarkers", c => xlSparklineStyle.MarkersColor = c);
        return xlSparklineStyle;

        void SetColor(string name, Action<XLColor> set)
        {
            if (slg.Element(SpreadsheetXml.X14 + name) is { } color)
            {
                set(SpreadsheetXml.ReadColor(color));
            }
        }
    }

    private static XLSparklineMarkers LoadSparklineMarkers(XElement slg)
    {
        XLSparklineMarkers markers = XLSparklineMarkers.None;
        Add("markers", XLSparklineMarkers.Markers);
        Add("high", XLSparklineMarkers.HighPoint);
        Add("low", XLSparklineMarkers.LowPoint);
        Add("first", XLSparklineMarkers.FirstPoint);
        Add("last", XLSparklineMarkers.LastPoint);
        Add("negative", XLSparklineMarkers.NegativePoints);
        return markers;

        void Add(string attributeName, XLSparklineMarkers marker)
        {
            if (SpreadsheetXml.Bool(slg, attributeName) ?? false)
            {
                markers |= marker;
            }
        }
    }

    private static void LoadSparklines(XElement slg, XLSparklineGroup xlSparklineGroup)
    {
        foreach (
            XElement sparkline in slg.Descendants(SpreadsheetXml.X14 + "sparklines")
                .SelectMany(sparklines => sparklines.Descendants(SpreadsheetXml.X14 + "sparkline"))
        )
        {
            // The sqlref must contain exactly one ref [MS-XLSX]. Excel ignores everything after the first one.
            string refText = (sparkline.Element(SpreadsheetXml.Xm + "sqref")?.Value ?? string.Empty)
                .Trim()
                .Split(' ')[0];
            Point location = Point.Parse(refText);

            // Technically, there could be more than one sparkline per cell, so use Set instead of Add.
            xlSparklineGroup.SetSparkline(
                location,
                sparkline.Element(SpreadsheetXml.Xm + "f")?.Value
            );
        }
    }

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

    #region Attributes of the streamed elements

    private static bool? Bool(XmlReader reader, string name)
    {
        string value = reader.GetAttribute(name);
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        if (value == "1" || string.Equals("true", value, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (value == "0" || string.Equals("false", value, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        throw new FormatException($"Unable to parse '{value}' to bool.");
    }

    private static int? Int(XmlReader reader, string name) =>
        reader.GetAttribute(name) is { Length: > 0 } value
            ? int.Parse(value, CultureInfo.InvariantCulture)
            : null;

    private static uint? UInt(XmlReader reader, string name) =>
        reader.GetAttribute(name) is { Length: > 0 } value
            ? uint.Parse(value, CultureInfo.InvariantCulture)
            : null;

    private static double? Double(XmlReader reader, string name) =>
        reader.GetAttribute(name) is { Length: > 0 } value
            ? double.Parse(value, NumberStyles.Float, XlsxSharp.XLHelper.ParseCulture)
            : null;

    private static double? Double(XmlReader reader, string name, string namespaceUri) =>
        reader.GetAttribute(name, namespaceUri) is { Length: > 0 } value
            ? double.Parse(value, NumberStyles.Float, XlsxSharp.XLHelper.ParseCulture)
            : null;

    /// <summary>
    /// An attribute of type <c>ST_CellRef</c>.
    /// </summary>
    private static Point? CellRef(XmlReader reader, string name) =>
        reader.GetAttribute(name) is { Length: > 0 } value ? Point.Parse(value) : null;

    /// <summary>
    /// An attribute of type <c>ST_Ref</c>.
    /// </summary>
    private static Area? Ref(XmlReader reader, string name) =>
        reader.GetAttribute(name) is { Length: > 0 } value ? Area.Parse(value) : null;

    #endregion
}

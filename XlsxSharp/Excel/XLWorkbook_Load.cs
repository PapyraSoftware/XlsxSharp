#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.Comments;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Drawings.Style;
using XlsxSharp.Excel.IO;
using XlsxSharp.Excel.Misc;
using XlsxSharp.Excel.Protection;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;
using XlsxSharp.IO;
using XlsxSharp.Utils;
using Run = DocumentFormat.OpenXml.Spreadsheet.Run;
using RunProperties = DocumentFormat.OpenXml.Spreadsheet.RunProperties;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XlsxSharp.Excel;

public partial class XLWorkbook
{
    private void Load(string file) => this.LoadSheets(file);

    private void Load(Stream stream) => this.LoadSheets(stream);

    private void LoadSheets(string fileName)
    {
        using (SpreadsheetDocument dSpreadsheet = SpreadsheetDocument.Open(fileName, false))
        {
            this.LoadSpreadsheetDocument(dSpreadsheet);
        }
    }

    private void LoadSheets(Stream stream)
    {
        using (SpreadsheetDocument dSpreadsheet = SpreadsheetDocument.Open(stream, false))
        {
            this.LoadSpreadsheetDocument(dSpreadsheet);
        }
    }

    private void LoadSheetsFromTemplate(string fileName)
    {
        using (SpreadsheetDocument dSpreadsheet = SpreadsheetDocument.CreateFromTemplate(fileName))
        {
            this.LoadSpreadsheetDocument(dSpreadsheet);
        }

        // If we load a workbook as a template, we have to treat it as a "new" workbook.
        // The original file will NOT be copied into place before changes are applied
        // Hence all loaded RelIds have to be cleared
        this.ResetAllRelIds();
    }

    private void ResetAllRelIds()
    {
        foreach (XLPivotCache pc in this.PivotCachesInternal)
        {
            pc.WorkbookCacheRelId = null;
        }

        uint sheetId = 1u;
        foreach (XLWorksheet ws in this.WorksheetsInternal)
        {
            // Ensure unique sheetId for each sheet.
            ws.SheetId = sheetId++;
            ws.RelId = null;

            foreach (XLPivotTable pt in ws.PivotTables.Cast<XLPivotTable>())
            {
                pt.CacheDefinitionRelId = null;
                pt.RelId = null;
            }

            foreach (XLPicture picture in ws.Pictures.Cast<XLPicture>())
            {
                picture.RelId = null;
            }

            foreach (XLTable table in ws.Tables.Cast<XLTable>())
            {
                table.RelId = null;
            }
        }
    }

    private void LoadSpreadsheetDocument(SpreadsheetDocument dSpreadsheet)
    {
        LoadContext context = new();
        this.ShapeIdManager = new XLIdManager();
        this.SetProperties(dSpreadsheet);

        SharedStringItem[] sharedStrings = null;
        WorkbookPart workbookPart = dSpreadsheet.WorkbookPart;
        if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any())
        {
            SharedStringTablePart shareStringPart = workbookPart
                .GetPartsOfType<SharedStringTablePart>()
                .First();
            sharedStrings = [.. shareStringPart.SharedStringTable.Elements<SharedStringItem>()];
        }

        LoadWorkbookTheme(workbookPart?.ThemePart, this);

        if (dSpreadsheet.CustomFilePropertiesPart != null)
        {
            foreach (
                CustomDocumentProperty m in dSpreadsheet.CustomFilePropertiesPart.Properties.Elements<CustomDocumentProperty>()
            )
            {
                string name = m.Name?.Value;

                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (m.VTLPWSTR != null)
                {
                    this.CustomProperties.Add(name, m.VTLPWSTR.Text);
                }
                else if (m.VTFileTime != null)
                {
                    this.CustomProperties.Add(
                        name,
                        DateTime.ParseExact(
                            m.VTFileTime.Text,
                            "yyyy'-'MM'-'dd'T'HH':'mm':'ssK",
                            CultureInfo.InvariantCulture
                        )
                    );
                }
                else if (m.VTDouble != null)
                {
                    this.CustomProperties.Add(
                        name,
                        double.Parse(m.VTDouble.Text, CultureInfo.InvariantCulture)
                    );
                }
                else if (m.VTBool != null)
                {
                    this.CustomProperties.Add(name, m.VTBool.Text == "true");
                }
            }
        }

        WorkbookProperties wbProps = workbookPart.Workbook.WorkbookProperties;
        if (wbProps != null)
        {
            this.Use1904DateSystem = OpenXmlHelper.GetBooleanValueAsBool(wbProps.Date1904, false);
        }

        FileSharing wbFilesharing = workbookPart.Workbook.FileSharing;
        if (wbFilesharing != null)
        {
            this.FileSharing.ReadOnlyRecommended = OpenXmlHelper.GetBooleanValueAsBool(
                wbFilesharing.ReadOnlyRecommended,
                false
            );
            this.FileSharing.UserName = wbFilesharing.UserName?.Value;
        }

        LoadWorkbookProtection(workbookPart.Workbook.WorkbookProtection, this);

        CalculationProperties calculationProperties = workbookPart.Workbook.CalculationProperties;
        if (calculationProperties != null)
        {
            EnumValue<CalculateModeValues> calculateMode = calculationProperties.CalculationMode;
            if (calculateMode != null)
            {
                this.CalculateMode = calculateMode.Value.ToClosedXml();
            }

            BooleanValue calculationOnSave = calculationProperties.CalculationOnSave;
            if (calculationOnSave != null)
            {
                this.CalculationOnSave = calculationOnSave.Value;
            }

            BooleanValue forceFullCalculation = calculationProperties.ForceFullCalculation;
            if (forceFullCalculation != null)
            {
                this.ForceFullCalculation = forceFullCalculation.Value;
            }

            BooleanValue fullCalculationOnLoad = calculationProperties.FullCalculationOnLoad;
            if (fullCalculationOnLoad != null)
            {
                this.FullCalculationOnLoad = fullCalculationOnLoad.Value;
            }

            BooleanValue fullPrecision = calculationProperties.FullPrecision;
            if (fullPrecision != null)
            {
                this.FullPrecision = fullPrecision.Value;
            }

            EnumValue<ReferenceModeValues> referenceMode = calculationProperties.ReferenceMode;
            if (referenceMode != null)
            {
                this.ReferenceStyle = referenceMode.Value.ToClosedXml();
            }
        }

        ExtendedFilePropertiesPart efp = dSpreadsheet.ExtendedFilePropertiesPart;
        if (efp != null && efp.Properties != null)
        {
            if (efp.Properties.Elements<Company>().Any())
            {
                this.Properties.Company = efp.Properties.GetFirstChild<Company>().Text;
            }

            if (efp.Properties.Elements<Manager>().Any())
            {
                this.Properties.Manager = efp.Properties.GetFirstChild<Manager>().Text;
            }
        }

        WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart is not null)
        {
            using XmlTreeReader xmlReader = this.CreateTreeReader(stylesPart);
            StylesReader stylesReader = new(xmlReader, this.Styles);
            stylesReader.Load();
        }

        // TODO Styles: Verify the column width is same as DefaultColumnWidth even if normal style is missing
        this.ColumnWidth = XlsxSharp.XLHelper.CalculateColumnWidth(8, this.Format.Font, this);

        // We loop through the sheets in 2 passes: first just to add the sheets and second to add all the data for the sheets.
        // We do this mainly because it skips a very costly calculation invalidation step, but it also make things more consistent,
        // e.g. when reading calculations that reference other sheets, we know that those sheets always already exist.
        // That consistency point isn't required yet but could be taken advantage of in the future.
        Sheets sheets = workbookPart.Workbook.Sheets;
        int position = 0;
        foreach (Sheet dSheet in sheets.OfType<Sheet>())
        {
            position++;
            StringValue sheetName = dSheet.Name;
            uint sheetId = dSheet.SheetId.Value;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                XLWorksheet emptySheet = this.WorksheetsInternal.Add(sheetName, position, sheetId);
                if (dSheet.State != null)
                {
                    emptySheet.Visibility = dSheet.State.Value.ToClosedXml();
                }

                continue;
            }

            // Although relationship to worksheet is most common, there can be other types
            // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            WorksheetPart worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
            if (worksheetPart == null)
            {
                this.UnsupportedSheets.Add(
                    new UnsupportedSheet { SheetId = sheetId, Position = position }
                );
                continue;
            }

            XLWorksheet ws = this.WorksheetsInternal.Add(sheetName, position, sheetId);
            ws.RelId = dSheet.Id;

            if (dSheet.State != null)
            {
                ws.Visibility = dSheet.State.Value.ToClosedXml();
            }
        }

        position = 0;
        foreach (Sheet dSheet in sheets.OfType<Sheet>())
        {
            position++;
            StringValue sheetName = dSheet.Name;
            uint sheetId = dSheet.SheetId.Value;

            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // Although relationship to worksheet is most common, there can be other types
            // than worksheet, e.g. chartSheet. Since we can't load them, add them to list
            // of unsupported sheets and copy them when saving. See Codeplex #6932.
            WorksheetPart worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;
            if (worksheetPart == null)
            {
                continue;
            }

            if (!this.WorksheetsInternal.TryGetWorksheet(sheetName, out XLWorksheet ws))
            {
                // This shouldn't be possible, as all worksheets should have already been added in the loop before this loop
                continue;
            }

            WorksheetPartReader worksheetPartReader = new();
            worksheetPartReader.LoadWorksheet(ws, worksheetPart, sharedStrings, context);

            ws.ConditionalFormats.ReorderAccordingToOriginalPriority();

            #region LoadTables

            foreach (TableDefinitionPart tableDefinitionPart in worksheetPart.TableDefinitionParts)
            {
                string relId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                Table dTable = tableDefinitionPart.Table;

                string reference = dTable.Reference.Value;
                string tableName = dTable.Name ?? dTable.DisplayName ?? string.Empty;
                if (string.IsNullOrWhiteSpace(tableName))
                {
                    throw new InvalidDataException("The table name is missing.");
                }

                XLTable xlTable = ws.Range(reference).CreateTable(tableName, false) as XLTable;
                xlTable.RelId = relId;

                // Add columns to the table
                foreach (TableColumn tableColumn in dTable.TableColumns)
                {
                    string fieldName = GetTableColumnName(tableColumn.Name.Value);
                    XLTableField xlField = xlTable.AddField(fieldName);

                    if (tableColumn.HeaderRowDifferentialFormattingId is { } headerDxfId)
                    {
                        xlField.HeaderFormatValue = this.Styles.DifferentialFormats[
                            checked((int)headerDxfId.Value)
                        ];
                    }

                    if (tableColumn.DataFormatId is { } dataDxfId)
                    {
                        xlField.DataFormatValue = this.Styles.DifferentialFormats[
                            checked((int)dataDxfId.Value)
                        ];
                    }

                    if (tableColumn.TotalsRowDifferentialFormattingId is { } totalsDxfId)
                    {
                        xlField.TotalFormatValue = this.Styles.DifferentialFormats[
                            checked((int)totalsDxfId.Value)
                        ];
                    }
                }

                if (dTable.HeaderRowCount != null && dTable.HeaderRowCount == 0)
                {
                    xlTable._showHeaderRow = false;
                }
                else
                {
                    xlTable.InitializeAutoFilter();
                }

                if (dTable.TotalsRowCount != null && dTable.TotalsRowCount.Value > 0)
                {
                    xlTable._showTotalsRow = true;
                }

                if (dTable.TableStyleInfo != null)
                {
                    if (dTable.TableStyleInfo.ShowFirstColumn != null)
                    {
                        xlTable.EmphasizeFirstColumn = dTable.TableStyleInfo.ShowFirstColumn.Value;
                    }

                    if (dTable.TableStyleInfo.ShowLastColumn != null)
                    {
                        xlTable.EmphasizeLastColumn = dTable.TableStyleInfo.ShowLastColumn.Value;
                    }

                    if (dTable.TableStyleInfo.ShowRowStripes != null)
                    {
                        xlTable.ShowRowStripes = dTable.TableStyleInfo.ShowRowStripes.Value;
                    }

                    if (dTable.TableStyleInfo.ShowColumnStripes != null)
                    {
                        xlTable.ShowColumnStripes = dTable.TableStyleInfo.ShowColumnStripes.Value;
                    }

                    if (dTable.TableStyleInfo.Name != null)
                    {
                        XLTableTheme theme = XLTableTheme.FromName(
                            dTable.TableStyleInfo.Name.Value
                        );
                        if (theme != null)
                        {
                            xlTable.Theme = theme;
                        }
                        else
                        {
                            xlTable.Theme = new XLTableTheme(dTable.TableStyleInfo.Name.Value);
                        }
                    }
                    else
                    {
                        xlTable.Theme = XLTableTheme.None;
                    }
                }

                if (dTable.AutoFilter != null)
                {
                    xlTable.ShowAutoFilter = true;
                    AutoFilterReader.LoadAutoFilterColumns(dTable.AutoFilter, xlTable.AutoFilter);
                }
                else
                {
                    xlTable.ShowAutoFilter = false;
                }

                if (xlTable.ShowTotalsRow)
                {
                    foreach (TableColumn tableColumn in dTable.TableColumns.Cast<TableColumn>())
                    {
                        string tableColumnName = GetTableColumnName(tableColumn.Name.Value);
                        if (tableColumn.TotalsRowFunction != null)
                        {
                            xlTable.Field(tableColumnName).TotalsRowFunction =
                                tableColumn.TotalsRowFunction.Value.ToClosedXml();
                        }

                        if (tableColumn.TotalsRowFormula != null)
                        {
                            xlTable.Field(tableColumnName).TotalsRowFormulaA1 = tableColumn
                                .TotalsRowFormula
                                .Text;
                        }

                        if (tableColumn.TotalsRowLabel != null)
                        {
                            xlTable.Field(tableColumnName).TotalsRowLabel = tableColumn
                                .TotalsRowLabel
                                .Value;
                        }
                    }
                    if (xlTable.AutoFilter != null)
                    {
                        xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                            xlTable.RangeAddress.FirstAddress.RowNumber,
                            xlTable.RangeAddress.FirstAddress.ColumnNumber,
                            xlTable.RangeAddress.LastAddress.RowNumber - 1,
                            xlTable.RangeAddress.LastAddress.ColumnNumber
                        );
                    }
                }
                else if (xlTable.AutoFilter != null)
                {
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);
                }
            }

            #endregion LoadTables

            LoadDrawings(worksheetPart, ws);

            #region LoadComments

            if (worksheetPart.WorksheetCommentsPart != null)
            {
                DocumentFormat.OpenXml.Spreadsheet.Comments root = worksheetPart
                    .WorksheetCommentsPart
                    .Comments;
                List<Author> authors =
                [
                    .. root.GetFirstChild<Authors>().ChildElements.OfType<Author>(),
                ];
                List<Comment> comments =
                [
                    .. root.GetFirstChild<CommentList>().ChildElements.OfType<Comment>(),
                ];

                // **** MAYBE FUTURE SHAPE SIZE SUPPORT
                IList<XElement> shapes = GetCommentShapes(worksheetPart);

                for (int i = 0; i < comments.Count; i++)
                {
                    Comment c = comments[i];

                    XElement shape = null;
                    if (i < shapes.Count)
                    {
                        shape = shapes[i];
                    }

                    // find cell by reference
                    XLCell cell = ws.Cell(c.Reference);

                    string shapeIdString = shape?.Attribute("id")?.Value;
                    if (shapeIdString?.StartsWith("_x0000_s") ?? false)
                    {
                        shapeIdString = shapeIdString.Substring(8);
                    }

                    int? shapeId = int.TryParse(shapeIdString, out int sid) ? (int?)sid : null;
                    XLComment xlComment = cell.CreateComment(shapeId);

                    xlComment.Author = authors[(int)c.AuthorId.Value].InnerText;
                    this.ShapeIdManager.Add(xlComment.ShapeId);

                    IEnumerable<Run> runs = c.GetFirstChild<CommentText>().Elements<Run>();
                    foreach (Run run in runs)
                    {
                        RunProperties runProperties = run.RunProperties;
                        string text = run.Text.InnerText.FixNewLines();
                        IXLRichString rt = xlComment.AddText(text);
                        OpenXmlHelper.LoadFont(runProperties, rt);
                    }

                    if (shape != null)
                    {
                        this.LoadShapeProperties(xlComment, shape);

                        XElement clientData = shape
                            .Elements()
                            .First(e => e.Name.LocalName == "ClientData");
                        this.LoadClientData(xlComment, clientData);

                        XElement textBox = shape
                            .Elements()
                            .First(e => e.Name.LocalName == "textbox");
                        this.LoadTextBox(xlComment, textBox);

                        XAttribute alt = shape.Attribute("alt");
                        if (alt != null)
                        {
                            xlComment.Style.Web.SetAlternateText(alt.Value);
                        }

                        this.LoadColorsAndLines(xlComment, shape);

                        //var insetmode = (string)shape.Attributes().First(a=> a.Name.LocalName == "insetmode");
                        //xlComment.Style.Margins.Automatic = insetmode != null && insetmode.Equals("auto");
                    }
                }
            }

            #endregion LoadComments
        }

        Workbook workbook = workbookPart.Workbook;

        BookViews bookViews = workbook.BookViews;
        if (bookViews != null && bookViews.FirstOrDefault() is WorkbookView workbookView)
        {
            if (workbookView.ActiveTab == null || !workbookView.ActiveTab.HasValue)
            {
                this.Worksheets.First().SetTabActive().Unhide();
            }
            else
            {
                UnsupportedSheet unsupportedSheet = this.UnsupportedSheets.FirstOrDefault(us =>
                    us.Position == (int)(workbookView.ActiveTab.Value + 1)
                );
                if (unsupportedSheet != null)
                {
                    unsupportedSheet.IsActive = true;
                }
                else
                {
                    this.Worksheet((int)(workbookView.ActiveTab.Value + 1)).SetTabActive();
                }
            }
        }
        this.LoadDefinedNames(workbook);

        // Read cache definition before table definition
        foreach (
            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart in workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>()
        )
        {
            XLPivotCache pivotCache = PivotTableCacheDefinitionPartReader.Load(
                workbookPart,
                pivotTableCacheDefinitionPart,
                this
            );
            if (pivotTableCacheDefinitionPart.PivotTableCacheRecordsPart is { } recordsPart)
            {
                using XmlTreeReader reader = this.CreateTreeReader(recordsPart);
                PivotCacheRecordsReader recordsReader = new(reader, pivotCache);
                recordsReader.ReadRecordsToCache();
            }
        }

        // Delay loading of pivot tables until all sheets have been loaded
        foreach (Sheet dSheet in sheets.OfType<Sheet>())
        {
            if (string.IsNullOrEmpty(dSheet.Id))
            {
                // Some non-Excel producers create sheets with empty relId.
                continue;
            }

            // The referenced sheet can also be ChartsheetPart. Only look for pivot tables in normal sheet parts.
            WorksheetPart worksheetPart = workbookPart.GetPartById(dSheet.Id) as WorksheetPart;

            if (worksheetPart is not null)
            {
                XLWorksheet ws = (XLWorksheet)this.WorksheetsInternal.Worksheet(dSheet.Name);

                foreach (PivotTablePart pivotTablePart in worksheetPart.PivotTableParts)
                {
                    PivotTableDefinitionPartReader.Load(
                        workbookPart,
                        pivotTablePart,
                        worksheetPart,
                        ws,
                        context
                    );
                }
            }
        }
    }

    private static void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws)
    {
        if (wsPart.DrawingsPart != null)
        {
            DrawingsPart drawingsPart = wsPart.DrawingsPart;

            foreach (OpenXmlElement anchor in drawingsPart.WorksheetDrawing.ChildElements)
            {
                string imgId = GetImageRelIdFromAnchor(anchor);

                //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
                if (imgId == null)
                {
                    continue;
                }

                OpenXmlPart imagePart = drawingsPart.GetPartById(imgId);
                using (Stream stream = imagePart.GetStream())
                using (MemoryStream ms = new())
                {
                    stream.CopyTo(ms);
                    Xdr.NonVisualDrawingProperties vsdp = GetPropertiesFromAnchor(anchor);

                    XLPicture picture =
                        ws.AddPicture(ms, vsdp.Name, Convert.ToInt32(vsdp.Id.Value)) as XLPicture;
                    picture.RelId = imgId;

                    Xdr.ShapeProperties spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                    picture.Placement = XLPicturePlacement.FreeFloating;

                    if (spPr?.Transform2D?.Extents?.Cx.HasValue ?? false)
                    {
                        picture.Width = ConvertFromEnglishMetricUnits(
                            spPr.Transform2D.Extents.Cx,
                            ws.Workbook.DpiX
                        );
                    }

                    if (spPr?.Transform2D?.Extents?.Cy.HasValue ?? false)
                    {
                        picture.Height = ConvertFromEnglishMetricUnits(
                            spPr.Transform2D.Extents.Cy,
                            ws.Workbook.DpiY
                        );
                    }

                    if (anchor is Xdr.AbsoluteAnchor)
                    {
                        Xdr.AbsoluteAnchor absoluteAnchor = anchor as Xdr.AbsoluteAnchor;
                        picture.MoveTo(
                            ConvertFromEnglishMetricUnits(
                                absoluteAnchor.Position.X.Value,
                                ws.Workbook.DpiX
                            ),
                            ConvertFromEnglishMetricUnits(
                                absoluteAnchor.Position.Y.Value,
                                ws.Workbook.DpiY
                            )
                        );
                    }
                    else if (anchor is Xdr.OneCellAnchor)
                    {
                        Xdr.OneCellAnchor oneCellAnchor = anchor as Xdr.OneCellAnchor;
                        XLMarker from = LoadMarker(ws, oneCellAnchor.FromMarker);
                        picture.MoveTo(from.Cell, from.Offset);
                    }
                    else if (anchor is Xdr.TwoCellAnchor)
                    {
                        Xdr.TwoCellAnchor twoCellAnchor = anchor as Xdr.TwoCellAnchor;
                        XLMarker from = LoadMarker(ws, twoCellAnchor.FromMarker);
                        XLMarker to = LoadMarker(ws, twoCellAnchor.ToMarker);

                        if (
                            twoCellAnchor.EditAs == null
                            || !twoCellAnchor.EditAs.HasValue
                            || twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell
                        )
                        {
                            picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                        }
                        else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                        {
                            Xdr.ShapeProperties shapeProperties = twoCellAnchor
                                .Descendants<Xdr.ShapeProperties>()
                                .FirstOrDefault();
                            if (shapeProperties != null)
                            {
                                picture.MoveTo(
                                    ConvertFromEnglishMetricUnits(
                                        spPr.Transform2D.Offset.X,
                                        ws.Workbook.DpiX
                                    ),
                                    ConvertFromEnglishMetricUnits(
                                        spPr.Transform2D.Offset.Y,
                                        ws.Workbook.DpiY
                                    )
                                );
                            }
                        }
                        else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                        {
                            picture.MoveTo(from.Cell, from.Offset);
                        }
                    }
                }
            }
        }
    }

    private static int ConvertFromEnglishMetricUnits(long emu, double resolution) =>
        Convert.ToInt32(emu * resolution / 914400);

    private static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
    {
        int row = Math.Min(
            XlsxSharp.XLHelper.MaxRowNumber,
            Math.Max(1, Convert.ToInt32(marker.RowId.InnerText) + 1)
        );
        int column = Math.Min(
            XlsxSharp.XLHelper.MaxColumnNumber,
            Math.Max(1, Convert.ToInt32(marker.ColumnId.InnerText) + 1)
        );
        return new XLMarker(
            ws.Cell(row, column),
            new System.Drawing.Point(
                ConvertFromEnglishMetricUnits(
                    Convert.ToInt32(marker.ColumnOffset.InnerText),
                    ws.Workbook.DpiX
                ),
                ConvertFromEnglishMetricUnits(
                    Convert.ToInt32(marker.RowOffset.InnerText),
                    ws.Workbook.DpiY
                )
            )
        );
    }

    #region Comment Helpers

    private static IList<XElement> GetCommentShapes(WorksheetPart worksheetPart)
    {
        // Cannot get this to return Vml.Shape elements
        foreach (VmlDrawingPart vmlPart in worksheetPart.VmlDrawingParts)
        {
            using (Stream stream = vmlPart.GetStream(FileMode.Open))
            {
                XDocument xdoc = XDocumentExtensions.Load(stream);
                if (xdoc == null)
                {
                    continue;
                }

                XElement root = xdoc.Root.Element("xml") ?? xdoc.Root;

                if (root == null)
                {
                    continue;
                }

                List<XElement> shapes =
                [
                    .. root.Elements(XName.Get("shape", "urn:schemas-microsoft-com:vml"))
                        .Where(e =>
                            new[]
                            {
                                "#" + XLConstants.Comment.ShapeTypeId,
                                "#" + XLConstants.Comment.AlternateShapeTypeId,
                            }.Contains(e.Attribute("type")?.Value)
                        ),
                ];

                if (shapes != null)
                {
                    return shapes;
                }
            }
        }

        throw new ArgumentException("Could not load comments file");
    }

    #endregion Comment Helpers

    private static string GetTableColumnName(string name) =>
        name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");

    private void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
    {
        XAttribute strokeColor = shape.Attribute(@"strokecolor");
        if (strokeColor is not null)
        {
            drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);
        }

        XAttribute strokeWeight = shape.Attribute(@"strokeweight");
        if (strokeWeight != null && this.TryGetPtValue(strokeWeight.Value, out double lineWeight))
        {
            drawing.Style.ColorsAndLines.LineWeight = lineWeight;
        }

        XAttribute fillColor = shape.Attribute(@"fillcolor");
        if (fillColor is not null)
        {
            drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);
        }

        XElement fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
        if (fill != null)
        {
            XAttribute opacity = fill.Attribute("opacity");
            if (opacity != null)
            {
                string opacityVal = opacity.Value;
                if (opacityVal.EndsWith("f"))
                {
                    drawing.Style.ColorsAndLines.FillTransparency =
                        double.Parse(
                            opacityVal.Substring(0, opacityVal.Length - 1),
                            CultureInfo.InvariantCulture
                        ) / 65536.0;
                }
                else
                {
                    drawing.Style.ColorsAndLines.FillTransparency = double.Parse(
                        opacityVal,
                        CultureInfo.InvariantCulture
                    );
                }
            }
        }

        XElement stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
        if (stroke != null)
        {
            XAttribute opacity = stroke.Attribute("opacity");
            if (opacity != null)
            {
                string opacityVal = opacity.Value;
                if (opacityVal.EndsWith("f"))
                {
                    drawing.Style.ColorsAndLines.LineTransparency =
                        double.Parse(
                            opacityVal.Substring(0, opacityVal.Length - 1),
                            CultureInfo.InvariantCulture
                        ) / 65536.0;
                }
                else
                {
                    drawing.Style.ColorsAndLines.LineTransparency = double.Parse(
                        opacityVal,
                        CultureInfo.InvariantCulture
                    );
                }
            }

            XAttribute dashStyle = stroke.Attribute("dashstyle");
            if (dashStyle != null)
            {
                string dashStyleVal = dashStyle.Value.ToLower();
                if (dashStyleVal == "1 1" || dashStyleVal == "shortdot")
                {
                    XAttribute endCap = stroke.Attribute("endcap");
                    if (endCap != null && endCap.Value == "round")
                    {
                        drawing.Style.ColorsAndLines.LineDash = XLDashStyle.RoundDot;
                    }
                    else
                    {
                        drawing.Style.ColorsAndLines.LineDash = XLDashStyle.SquareDot;
                    }
                }
                else
                {
                    switch (dashStyleVal)
                    {
                        case "dash":
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.Dash;
                            break;
                        case "dashdot":
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.DashDot;
                            break;
                        case "longdash":
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDash;
                            break;
                        case "longdashdot":
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDot;
                            break;
                        case "longdashdotdot":
                            drawing.Style.ColorsAndLines.LineDash = XLDashStyle.LongDashDotDot;
                            break;
                    }
                }
            }

            XAttribute lineStyle = stroke.Attribute("linestyle");
            if (lineStyle != null)
            {
                string lineStyleVal = lineStyle.Value.ToLower();
                switch (lineStyleVal)
                {
                    case "single":
                        drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.Single;
                        break;
                    case "thickbetweenthin":
                        drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickBetweenThin;
                        break;
                    case "thickthin":
                        drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThickThin;
                        break;
                    case "thinthick":
                        drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThick;
                        break;
                    case "thinthin":
                        drawing.Style.ColorsAndLines.LineStyle = XLLineStyle.ThinThin;
                        break;
                }
            }
        }
    }

    private void LoadTextBox<T>(IXLDrawing<T> xlDrawing, XElement textBox)
    {
        XAttribute attStyle = textBox.Attribute("style");
        if (attStyle != null)
        {
            LoadTextBoxStyle(xlDrawing, attStyle);
        }

        XAttribute attInset = textBox.Attribute("inset");
        if (attInset != null)
        {
            this.LoadTextBoxInset(xlDrawing, attInset);
        }
    }

    private void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset)
    {
        string[] split = attInset.Value.Split(',');
        xlDrawing.Style.Margins.Left = GetInsetInInches(split[0], this.DpiX);
        xlDrawing.Style.Margins.Top = GetInsetInInches(split[1], this.DpiY);
        xlDrawing.Style.Margins.Right = GetInsetInInches(split[2], this.DpiX);
        xlDrawing.Style.Margins.Bottom = GetInsetInInches(split[3], this.DpiY);
    }

    /// <summary>
    /// List of all VML length units and their conversion. Key is a name, value is a conversion
    /// function to EMU. See <a href="https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-units">documentation</a>.
    /// </summary>
    /// <remarks>
    /// OI-29500 says <em>Office also uses EMUs throughout VML as a valid unit system</em>.
    /// Relative units conversions are guesstimated by how Excel 2022 behaves for inset
    /// attribute of <c>TextBox</c> element of a note/comment. Generally speaking, Excel
    /// converts relative values to physical length (e.g. <c>px</c> to <c>pt</c>) and saves
    /// them as such. The <c>ex</c>/<c>em</c> units are not interpreted as described in the
    /// doc, but as 1/90th or an inch. The <c>%</c> seems to be always 0.
    /// </remarks>
    private static readonly Dictionary<string, Func<double, double, Emu?>> VmlLengthUnits = new()
    {
        { "in", (value, _) => Emu.From(value, AbsLengthUnit.Inch) },
        { "cm", (value, _) => Emu.From(value, AbsLengthUnit.Centimeter) },
        { "mm", (value, _) => Emu.From(value, AbsLengthUnit.Millimeter) },
        { "pt", (value, _) => Emu.From(value, AbsLengthUnit.Point) },
        { "pc", (value, _) => Emu.From(value, AbsLengthUnit.Pica) },
        { "emu", (value, _) => Emu.From(value, AbsLengthUnit.Emu) },
        { "px", (value, dpi) => Emu.From(value / dpi, AbsLengthUnit.Inch) },
        { "em", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        { "ex", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        { "%", (_, _) => Emu.ZeroPt },
    };

    private static double GetInsetInInches(string value, double dpi)
    {
        string unit = value.Trim();
        foreach ((string unitName, Func<double, double, Emu?> conversion) in VmlLengthUnits)
        {
            if (
                unit.EndsWith(unitName)
                && double.TryParse(
                    unit[..^unitName.Length],
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double unitValue
                )
            )
            {
                Emu insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                return insetEmu.To(AbsLengthUnit.Inch);
            }
        }

        // Excel treats no/unexpected unit as 0
        return 0;
    }

    private static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
    {
        string style = attStyle.Value;
        string[] attributes = style.Split(';');
        foreach (string pair in attributes)
        {
            string[] split = pair.Split(':');
            if (split.Length != 2)
            {
                continue;
            }

            string attribute = split[0].Trim().ToLower();
            string value = split[1].Trim();
            bool isVertical = false;
            switch (attribute)
            {
                case "mso-fit-shape-to-text":
                    xlDrawing.Style.Size.SetAutomaticSize(value.Equals("t"));
                    break;
                case "mso-layout-flow-alt":
                    if (value.Equals("bottom-to-top"))
                    {
                        xlDrawing.Style.Alignment.SetOrientation(
                            XLDrawingTextOrientation.BottomToTop
                        );
                    }
                    else if (value.Equals("top-to-bottom"))
                    {
                        xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
                    }

                    break;

                case "layout-flow":
                    isVertical = value.Equals("vertical");
                    break;
                case "mso-direction-alt":
                    if (value == "auto")
                    {
                        xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.Context;
                    }

                    break;
                case "direction":
                    if (value == "RTL")
                    {
                        xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.RightToLeft;
                    }

                    break;
            }
            if (
                isVertical
                && xlDrawing.Style.Alignment.Orientation == XLDrawingTextOrientation.LeftToRight
            )
            {
                xlDrawing.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
            }
        }
    }

    private void LoadClientData<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        XElement anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
        if (anchor != null)
        {
            LoadClientDataAnchor<T>(drawing, anchor);
        }

        LoadDrawingPositioning<T>(drawing, clientData);
        LoadDrawingProtection(drawing, clientData);

        XElement visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
        drawing.Visible =
            visible != null
            && (
                string.IsNullOrEmpty(visible.Value)
                || visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase)
            );

        LoadDrawingHAlignment(drawing, clientData);
        LoadDrawingVAlignment(drawing, clientData);
    }

    private static void LoadDrawingHAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        XElement textHAlign = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
        if (textHAlign != null)
        {
            drawing.Style.Alignment.Horizontal = (XLDrawingHorizontalAlignment)
                Enum.Parse(typeof(XLDrawingHorizontalAlignment), textHAlign.Value.ToProper());
        }
    }

    private static void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        XElement textVAlign = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
        if (textVAlign != null)
        {
            drawing.Style.Alignment.Vertical = (XLDrawingVerticalAlignment)
                Enum.Parse(typeof(XLDrawingVerticalAlignment), textVAlign.Value.ToProper());
        }
    }

    private static void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        XElement lockedElement = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "Locked");
        XElement lockTextElement = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "LockText");
        bool locked = lockedElement != null && lockedElement.Value.ToLower() == "true";
        bool lockText = lockTextElement != null && lockTextElement.Value.ToLower() == "true";
        drawing.Style.Protection.Locked = locked;
        drawing.Style.Protection.LockText = lockText;
    }

    private static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        XElement moveWithCellsElement = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
        XElement sizeWithCellsElement = clientData
            .Elements()
            .FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
        bool moveWithCells = !(
            moveWithCellsElement != null && moveWithCellsElement.Value.ToLower() == "true"
        );
        bool sizeWithCells = !(
            sizeWithCellsElement != null && sizeWithCellsElement.Value.ToLower() == "true"
        );
        if (moveWithCells && !sizeWithCells)
        {
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
        }
        else if (moveWithCells && sizeWithCells)
        {
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
        }
        else
        {
            drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
        }
    }

    private static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
    {
        string[] location = anchor.Value.Split(',');
        drawing.Position.Column = int.Parse(location[0]) + 1;
        drawing.Position.ColumnOffset =
            double.Parse(location[1], CultureInfo.InvariantCulture) / 7.5;
        drawing.Position.Row = int.Parse(location[2]) + 1;
        drawing.Position.RowOffset = double.Parse(location[3], CultureInfo.InvariantCulture);
    }

    private void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
    {
        if (shape.Attribute("style") == null)
        {
            return;
        }

        foreach (string attributePair in shape.Attribute("style").Value.Split(';'))
        {
            string[] split = attributePair.Split(':');
            if (split.Length != 2)
            {
                continue;
            }

            string attribute = split[0].Trim().ToLower();
            string value = split[1].Trim();

            switch (attribute)
            {
                case "visibility":
                    xlDrawing.Visible = string.Equals(
                        "visible",
                        value,
                        StringComparison.OrdinalIgnoreCase
                    );
                    break;
                case "width":
                    if (this.TryGetPtValue(value, out double ptWidth))
                    {
                        xlDrawing.Style.Size.Width = ptWidth / 7.5;
                    }
                    break;

                case "height":
                    if (this.TryGetPtValue(value, out double ptHeight))
                    {
                        xlDrawing.Style.Size.Height = ptHeight;
                    }
                    break;

                case "z-index":
                    if (int.TryParse(value, out int zOrder))
                    {
                        xlDrawing.ZOrder = zOrder;
                    }
                    break;
            }
        }
    }

    private readonly Dictionary<string, double> knownUnits = new()
    {
        { "pt", 1.0 },
        { "in", 72.0 },
        { "mm", 72.0 / 25.4 },
    };

    private bool TryGetPtValue(string value, out double result)
    {
        KeyValuePair<string, double> knownUnit = this.knownUnits.FirstOrDefault(ku =>
            value.Contains(ku.Key)
        );

        if (knownUnit.Key == null)
        {
            return double.TryParse(value, out result);
        }

        value = value.Replace(knownUnit.Key, string.Empty);

        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        {
            result *= knownUnit.Value;
            return true;
        }

        result = 0d;
        return false;
    }

    private void LoadDefinedNames(Workbook workbook)
    {
        if (workbook.DefinedNames == null)
        {
            return;
        }

        foreach (DefinedName definedName in workbook.DefinedNames.OfType<DefinedName>())
        {
            StringValue name = definedName.Name;
            bool visible = true;
            if (definedName.Hidden != null)
            {
                visible = !BooleanValue.ToBoolean(definedName.Hidden);
            }

            int localSheetId = -1;
            if (definedName.LocalSheetId?.HasValue ?? false)
            {
                localSheetId = Convert.ToInt32(definedName.LocalSheetId.Value);
            }

            if (name == "_xlnm.Print_Area")
            {
                IEnumerable<string> fixedNames = validateDefinedNames(definedName.Text.Split(','));
                foreach (string area in fixedNames)
                {
                    if (area.Contains("["))
                    {
                        XLWorksheet ws = this.WorksheetsInternal.FirstOrDefault<XLWorksheet>(w =>
                            w.SheetId == (localSheetId + 1)
                        );
                        if (ws != null)
                        {
                            ws.PageSetup.PrintAreas.Add(area);
                        }
                    }
                    else
                    {
                        ParseReference(area, out string sheetName, out string sheetArea);
                        if (
                            !(
                                sheetArea.Equals("#REF")
                                || sheetArea.EndsWith("#REF!")
                                || sheetArea.Length == 0
                                || sheetName.Length == 0
                            )
                        )
                        {
                            this.WorksheetsInternal.Worksheet(sheetName)
                                .PageSetup.PrintAreas.Add(sheetArea);
                        }
                    }
                }
            }
            else if (name == "_xlnm.Print_Titles")
            {
                this.LoadPrintTitles(definedName);
            }
            else
            {
                string text = definedName.Text;

                StringValue comment = definedName.Comment;
                if (localSheetId == -1)
                {
                    if (this.DefinedNamesInternal.All<XLDefinedName>(nr => nr.Name != name))
                    {
                        this
                            .DefinedNamesInternal.Add(
                                name,
                                text,
                                comment,
                                validateName: false,
                                validateRangeAddress: false
                            )
                            .Visible = visible;
                    }
                }
                else
                {
                    if (this.Worksheet(localSheetId + 1).DefinedNames.All(nr => nr.Name != name))
                    {
                        ((XLDefinedNames)this.Worksheet(localSheetId + 1).DefinedNames)
                            .Add(
                                name,
                                text,
                                comment,
                                validateName: false,
                                validateRangeAddress: false
                            )
                            .Visible = visible;
                    }
                }
            }
        }
    }

    private static Regex definedNameRegex = new(@"\A('?).*\1!.*\z", RegexOptions.Compiled);

    private static IEnumerable<string> validateDefinedNames(IEnumerable<string> definedNames)
    {
        StringBuilder sb = new();
        foreach (string testName in definedNames)
        {
            if (sb.Length > 0)
            {
                sb.Append(',');
            }

            sb.Append(testName);

            Match matchedValidPattern = definedNameRegex.Match(sb.ToString());
            if (matchedValidPattern.Success)
            {
                yield return sb.ToString();
                sb = new StringBuilder();
            }
        }

        if (sb.Length > 0)
        {
            yield return sb.ToString();
        }
    }

    private void LoadPrintTitles(DefinedName definedName)
    {
        IEnumerable<string> areas = validateDefinedNames(definedName.Text.Split(','));
        foreach (string item in areas)
        {
            if (this.Range(item) != null)
            {
                this.SetColumnsOrRowsToRepeat(item);
            }
        }
    }

    private void SetColumnsOrRowsToRepeat(string area)
    {
        ParseReference(area, out string sheetName, out string sheetArea);
        sheetArea = sheetArea.Replace("$", "");

        if (sheetArea.Equals("#REF"))
        {
            return;
        }

        if (IsColReference(sheetArea))
        {
            this.WorksheetsInternal.Worksheet(sheetName)
                .PageSetup.SetColumnsToRepeatAtLeft(sheetArea);
        }

        if (IsRowReference(sheetArea))
        {
            this.WorksheetsInternal.Worksheet(sheetName).PageSetup.SetRowsToRepeatAtTop(sheetArea);
        }
    }

    // either $A:$X => true or $1:$99 => false
    private static bool IsColReference(string sheetArea) =>
        sheetArea.All(c => c == ':' || char.IsLetter(c));

    private static bool IsRowReference(string sheetArea) =>
        sheetArea.All(c => c == ':' || char.IsNumber(c));

    private static void ParseReference(string item, out string sheetName, out string sheetArea)
    {
        string[] sections = item.Trim().Split('!');
        if (sections.Length == 1)
        {
            sheetName = string.Empty;
            sheetArea = item;
        }
        else
        {
            sheetName = string.Join("!", sections.Take(sections.Length - 1)).UnescapeSheetName();
            sheetArea = sections[sections.Length - 1];
        }
    }

    private static void LoadWorkbookTheme(ThemePart tp, XLWorkbook wb)
    {
        if (tp is null)
        {
            return;
        }

        ColorScheme colorScheme = tp.Theme?.ThemeElements?.ColorScheme;
        if (colorScheme is not null)
        {
            string background1 = colorScheme.Light1Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(background1))
            {
                wb.Theme.Background1 = XLColor.FromHexRgb(background1);
            }
            string text1 = colorScheme.Dark1Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(text1))
            {
                wb.Theme.Text1 = XLColor.FromHexRgb(text1);
            }
            string background2 = colorScheme.Light2Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(background2))
            {
                wb.Theme.Background2 = XLColor.FromHexRgb(background2);
            }
            string text2 = colorScheme.Dark2Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(text2))
            {
                wb.Theme.Text2 = XLColor.FromHexRgb(text2);
            }
            string accent1 = colorScheme.Accent1Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent1))
            {
                wb.Theme.Accent1 = XLColor.FromHexRgb(accent1);
            }
            string accent2 = colorScheme.Accent2Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent2))
            {
                wb.Theme.Accent2 = XLColor.FromHexRgb(accent2);
            }
            string accent3 = colorScheme.Accent3Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent3))
            {
                wb.Theme.Accent3 = XLColor.FromHexRgb(accent3);
            }
            string accent4 = colorScheme.Accent4Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent4))
            {
                wb.Theme.Accent4 = XLColor.FromHexRgb(accent4);
            }
            string accent5 = colorScheme.Accent5Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent5))
            {
                wb.Theme.Accent5 = XLColor.FromHexRgb(accent5);
            }
            string accent6 = colorScheme.Accent6Color?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(accent6))
            {
                wb.Theme.Accent6 = XLColor.FromHexRgb(accent6);
            }
            string hyperlink = colorScheme.Hyperlink?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrEmpty(hyperlink))
            {
                wb.Theme.Hyperlink = XLColor.FromHexRgb(hyperlink);
            }
            string followedHyperlink = colorScheme
                .FollowedHyperlinkColor
                ?.RgbColorModelHex
                ?.Val
                ?.Value;
            if (!string.IsNullOrEmpty(followedHyperlink))
            {
                wb.Theme.FollowedHyperlink = XLColor.FromHexRgb(followedHyperlink);
            }
        }
    }

    private static void LoadWorkbookProtection(WorkbookProtection wp, XLWorkbook wb)
    {
        if (wp == null)
        {
            return;
        }

        wb.Protection.IsProtected = true;

        string algorithmName = wp.WorkbookAlgorithmName?.Value ?? string.Empty;
        if (string.IsNullOrEmpty(algorithmName))
        {
            wb.Protection.PasswordHash = wp.WorkbookPassword?.Value ?? string.Empty;
            wb.Protection.Base64EncodedSalt = string.Empty;
        }
        else if (
            DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.IsValidDescription(algorithmName)
        )
        {
            wb.Protection.Algorithm =
                DescribedEnumParser<XLProtectionAlgorithm.Algorithm>.FromDescription(algorithmName);
            wb.Protection.PasswordHash = wp.WorkbookHashValue?.Value ?? string.Empty;
            wb.Protection.SpinCount = wp.WorkbookSpinCount?.Value ?? 0;
            wb.Protection.Base64EncodedSalt = wp.WorkbookSaltValue?.Value ?? string.Empty;
        }

        wb.Protection.AllowElement(
            XLWorkbookProtectionElements.Structure,
            !OpenXmlHelper.GetBooleanValueAsBool(wp.LockStructure, false)
        );
        wb.Protection.AllowElement(
            XLWorkbookProtectionElements.Windows,
            !OpenXmlHelper.GetBooleanValueAsBool(wp.LockWindows, false)
        );
    }

    private void SetProperties(SpreadsheetDocument dSpreadsheet)
    {
        IPackageProperties p = dSpreadsheet.PackageProperties;
        this.Properties.Author = p.Creator;
        this.Properties.Category = p.Category;
        this.Properties.Comments = p.Description;
        if (p.Created != null)
        {
            this.Properties.Created = p.Created.Value;
        }

        if (p.Modified != null)
        {
            this.Properties.Modified = p.Modified.Value;
        }

        this.Properties.Keywords = p.Keywords;
        this.Properties.LastModifiedBy = p.LastModifiedBy;
        this.Properties.Status = p.ContentStatus;
        this.Properties.Subject = p.Subject;
        this.Properties.Title = p.Title;
    }

    private XmlTreeReader CreateTreeReader(OpenXmlPart openXmlPart)
    {
        Stream stream = openXmlPart.GetStream(FileMode.Open);
        return new XmlTreeReader(stream, XmlToEnumMapper.Instance, this.StrictAttributeParsing);
    }
}

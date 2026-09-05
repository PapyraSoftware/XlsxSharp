#nullable disable

using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.Excel.ContentManagers;
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
using XlsxSharp.Utils;
using static XlsxSharp.Excel.IO.OpenXmlConst;
using static XlsxSharp.Excel.XLWorkbook;
using Break = DocumentFormat.OpenXml.Spreadsheet.Break;
using Column = DocumentFormat.OpenXml.Spreadsheet.Column;
using Columns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Hyperlink = DocumentFormat.OpenXml.Spreadsheet.Hyperlink;
using OfficeExcel = DocumentFormat.OpenXml.Office.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XlsxSharp.Excel.IO;

internal class WorksheetPartWriter
{
    internal static void GenerateWorksheetPartContent(
        bool partIsEmpty,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveOptions options,
        SaveContext context
    )
    {
        Worksheet worksheetDom = GetWorksheetDom(
            partIsEmpty,
            worksheetPart,
            xlWorksheet,
            options,
            context
        );

        // The cells are written from the workbook model rather than from what the part held, so
        // the rows loaded above are dropped before the rest of the sheet is turned into XML.
        worksheetDom.GetFirstChild<SheetData>()?.RemoveAllChildren();

        XElement worksheet = ToXml(worksheetDom);
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

    private static Worksheet GetWorksheetDom(
        bool partIsEmpty,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveOptions options,
        SaveContext context
    )
    {
        if (options.ConsolidateConditionalFormatRanges)
        {
            xlWorksheet.ConditionalFormats.Consolidate();
        }

        #region Worksheet

        Worksheet worksheet;
        if (!partIsEmpty)
        {
            // Accessing the worksheet through worksheetPart.Worksheet creates an attached DOM
            // worksheet that is tracked and later saved automatically to the part.
            // Using the reader, we get a detached DOM.
            // The OpenXmlReader.Create method only reads xml declaration, but doesn't read content.
            using OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            if (!reader.Read())
            {
                throw new ArgumentException(
                    "Worksheet part should contain worksheet xml, but is empty."
                );
            }

            worksheet = (Worksheet)reader.LoadCurrentElement();
        }
        else
        {
            worksheet = new Worksheet();
        }

        if (worksheet.NamespaceDeclarations.All(ns => ns.Value != RelationshipsNs))
        {
            worksheet.AddNamespaceDeclaration("r", RelationshipsNs);
        }

        // We store the x14ac:dyDescent attribute (if set by a xlRow) in a row element. It's an optional attribute and it
        // needs a declared namespace. To avoid writing namespace to each <x:row> element during streaming, write it to
        // every sheet part ahead of time. The namespace has to be marked as ignorable, because OpenXML SDK validator will
        // refuse to validate it because it's an optional extension (see ISO29500 part 3).
        if (worksheet.NamespaceDeclarations.All(ns => ns.Value != X14Ac2009SsNs))
        {
            worksheet.AddNamespaceDeclaration("x14ac", X14Ac2009SsNs);
            worksheet.SetAttribute(
                new OpenXmlAttribute("mc", "Ignorable", MarkupCompatibilityNs, "x14ac")
            );
        }

        #endregion Worksheet

        XLWorksheetContentManager cm = new(worksheet);

        #region SheetProperties

        if (worksheet.SheetProperties == null)
        {
            worksheet.SheetProperties = new SheetProperties();
        }

        worksheet.SheetProperties.TabColor = xlWorksheet.TabColor.HasValue
            ? new TabColor().FromXlsxSharpColor<TabColor>(xlWorksheet.TabColor)
            : null;

        cm.SetElement(XLWorksheetContents.SheetProperties, worksheet.SheetProperties);

        if (worksheet.SheetProperties.OutlineProperties == null)
        {
            worksheet.SheetProperties.OutlineProperties = new OutlineProperties();
        }

        worksheet.SheetProperties.OutlineProperties.SummaryBelow = (
            xlWorksheet.Outline.SummaryVLocation == XLOutlineSummaryVLocation.Bottom
        );
        worksheet.SheetProperties.OutlineProperties.SummaryRight = (
            xlWorksheet.Outline.SummaryHLocation == XLOutlineSummaryHLocation.Right
        );

        if (
            worksheet.SheetProperties.PageSetupProperties == null
            && (xlWorksheet.PageSetup.PagesTall > 0 || xlWorksheet.PageSetup.PagesWide > 0)
        )
        {
            worksheet.SheetProperties.PageSetupProperties = new PageSetupProperties
            {
                FitToPage = true,
            };
        }

        #endregion SheetProperties

        // Empty worksheets have dimension A1 (not A1:A1)
        string sheetDimensionReference = "A1";
        if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
        {
            int maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
            int maxRow = xlWorksheet.Internals.CellsCollection.MaxRowUsed;
            sheetDimensionReference =
                "A1:"
                + XlsxSharp.XLHelper.GetColumnLetterFromNumber(maxColumn)
                + maxRow.ToInvariantString();
        }

        #region SheetViews

        if (worksheet.SheetDimension == null)
        {
            worksheet.SheetDimension = new SheetDimension { Reference = sheetDimensionReference };
        }

        cm.SetElement(XLWorksheetContents.SheetDimension, worksheet.SheetDimension);

        if (worksheet.SheetViews == null)
        {
            worksheet.SheetViews = new SheetViews();
        }

        cm.SetElement(XLWorksheetContents.SheetViews, worksheet.SheetViews);

        SheetView sheetView = (SheetView)worksheet.SheetViews.FirstOrDefault();
        if (sheetView == null)
        {
            sheetView = new SheetView { WorkbookViewId = 0U };
            worksheet.SheetViews.AppendChild(sheetView);
        }

        XLSheetViewContentManager svcm = new(sheetView);

        if (xlWorksheet.TabSelected)
        {
            sheetView.TabSelected = true;
        }
        else
        {
            sheetView.TabSelected = null;
        }

        if (xlWorksheet.RightToLeft)
        {
            sheetView.RightToLeft = true;
        }
        else
        {
            sheetView.RightToLeft = null;
        }

        if (xlWorksheet.ShowFormulas)
        {
            sheetView.ShowFormulas = true;
        }
        else
        {
            sheetView.ShowFormulas = null;
        }

        if (xlWorksheet.ShowGridLines)
        {
            sheetView.ShowGridLines = null;
        }
        else
        {
            sheetView.ShowGridLines = false;
        }

        if (xlWorksheet.ShowOutlineSymbols)
        {
            sheetView.ShowOutlineSymbols = null;
        }
        else
        {
            sheetView.ShowOutlineSymbols = false;
        }

        if (xlWorksheet.ShowRowColHeaders)
        {
            sheetView.ShowRowColHeaders = null;
        }
        else
        {
            sheetView.ShowRowColHeaders = false;
        }

        if (xlWorksheet.ShowRuler)
        {
            sheetView.ShowRuler = null;
        }
        else
        {
            sheetView.ShowRuler = false;
        }

        if (xlWorksheet.ShowWhiteSpace)
        {
            sheetView.ShowWhiteSpace = null;
        }
        else
        {
            sheetView.ShowWhiteSpace = false;
        }

        if (xlWorksheet.ShowZeros)
        {
            sheetView.ShowZeros = null;
        }
        else
        {
            sheetView.ShowZeros = false;
        }

        if (xlWorksheet.RightToLeft)
        {
            sheetView.RightToLeft = true;
        }
        else
        {
            sheetView.RightToLeft = null;
        }

        if (xlWorksheet.SheetView.View == XLSheetViewOptions.Normal)
        {
            sheetView.View = null;
        }
        else
        {
            sheetView.View = xlWorksheet.SheetView.View.ToOpenXml();
        }

        Pane pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (pane == null)
        {
            pane = new Pane();
            sheetView.InsertAt(pane, 0);
        }

        svcm.SetElement(XLSheetViewContents.Pane, pane);

        pane.State = PaneStateValues.FrozenSplit;
        int hSplit = xlWorksheet.SheetView.SplitColumn;
        int ySplit = xlWorksheet.SheetView.SplitRow;

        pane.HorizontalSplit = hSplit;
        pane.VerticalSplit = ySplit;

        // When panes are frozen, which part should move.
        PaneValues split;
        if (ySplit == 0 && hSplit == 0)
        {
            split = PaneValues.TopLeft;
        }
        else if (ySplit == 0 && hSplit != 0)
        {
            split = PaneValues.TopRight;
        }
        else if (ySplit != 0 && hSplit == 0)
        {
            split = PaneValues.BottomLeft;
        }
        else if (ySplit != 0 && hSplit != 0)
        {
            split = PaneValues.BottomRight;
        }

        pane.ActivePane = split;

        pane.TopLeftCell =
            XlsxSharp.XLHelper.GetColumnLetterFromNumber(xlWorksheet.SheetView.SplitColumn + 1)
            + (xlWorksheet.SheetView.SplitRow + 1);

        if (hSplit == 0 && ySplit == 0)
        {
            // We don't have a pane. Just a regular sheet.
            pane = null;
            sheetView.RemoveAllChildren<Pane>();
            svcm.SetElement(XLSheetViewContents.Pane, null);
        }

        // Do sheet view. Whether it's for a regular sheet or for the bottom-right pane
        if (
            !xlWorksheet.SheetView.TopLeftCellAddress.IsValid
            || xlWorksheet.SheetView.TopLeftCellAddress
                == new XLAddress(1, 1, fixedRow: false, fixedColumn: false)
        )
        {
            sheetView.TopLeftCell = null;
        }
        else
        {
            sheetView.TopLeftCell = xlWorksheet.SheetView.TopLeftCellAddress.ToString();
        }

        if (xlWorksheet.SelectedRanges.Any() || xlWorksheet.ActiveCell is not null)
        {
            sheetView.RemoveAllChildren<Selection>();
            svcm.SetElement(XLSheetViewContents.Selection, null);

            IXLRange firstSelection = xlWorksheet.SelectedRanges.FirstOrDefault();

            Action<Selection> populateSelection = (Selection selection) =>
            {
                if (xlWorksheet.ActiveCell is not null)
                {
                    selection.ActiveCell = xlWorksheet.ActiveCell.Value.ToString();
                }
                else if (firstSelection != null)
                {
                    selection.ActiveCell =
                        firstSelection.RangeAddress.FirstAddress.ToStringRelative(false);
                }

                List<string> seqRef =
                [
                    selection.ActiveCell.Value,
                    .. xlWorksheet.SelectedRanges.Select(range =>
                    {
                        if (range.RangeAddress.FirstAddress.Equals(range.RangeAddress.LastAddress))
                        {
                            return range.RangeAddress.FirstAddress.ToStringRelative(false);
                        }
                        else
                        {
                            return range.RangeAddress.ToStringRelative(false);
                        }
                    }),
                ];

                selection.SequenceOfReferences = new ListValue<StringValue>
                {
                    InnerText = string.Join(" ", seqRef.Distinct().ToArray()),
                };

                sheetView.InsertAfter(
                    selection,
                    svcm.GetPreviousElementFor(XLSheetViewContents.Selection)
                );
                svcm.SetElement(XLSheetViewContents.Selection, selection);
            };

            // If a pane exists, we need to set the active pane too
            // Yes, this might lead to 2 Selection elements!
            if (pane != null)
            {
                populateSelection(new Selection() { Pane = pane.ActivePane });
            }
            populateSelection(new Selection());
        }

        if (xlWorksheet.SheetView.ZoomScale == 100)
        {
            sheetView.ZoomScale = null;
        }
        else
        {
            sheetView.ZoomScale = (uint)
                Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScale));
        }

        if (xlWorksheet.SheetView.ZoomScaleNormal == 100)
        {
            sheetView.ZoomScaleNormal = null;
        }
        else
        {
            sheetView.ZoomScaleNormal = (uint)
                Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleNormal));
        }

        if (xlWorksheet.SheetView.ZoomScalePageLayoutView == 100)
        {
            sheetView.ZoomScalePageLayoutView = null;
        }
        else
        {
            sheetView.ZoomScalePageLayoutView = (uint)
                Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScalePageLayoutView));
        }

        if (xlWorksheet.SheetView.ZoomScaleSheetLayoutView == 100)
        {
            sheetView.ZoomScaleSheetLayoutView = null;
        }
        else
        {
            sheetView.ZoomScaleSheetLayoutView = (uint)
                Math.Max(10, Math.Min(400, xlWorksheet.SheetView.ZoomScaleSheetLayoutView));
        }

        #endregion SheetViews

        int maxOutlineColumn = 0;
        if (xlWorksheet.ColumnCount() > 0)
        {
            maxOutlineColumn = xlWorksheet.GetMaxColumnOutline();
        }

        int maxOutlineRow = 0;
        if (xlWorksheet.RowCount() > 0)
        {
            maxOutlineRow = xlWorksheet.GetMaxRowOutline();
        }

        #region SheetFormatProperties

        if (worksheet.SheetFormatProperties == null)
        {
            worksheet.SheetFormatProperties = new SheetFormatProperties();
        }

        cm.SetElement(XLWorksheetContents.SheetFormatProperties, worksheet.SheetFormatProperties);

        worksheet.SheetFormatProperties.DefaultRowHeight = xlWorksheet.RowHeight.SaveRound();

        if (xlWorksheet.RowHeightChanged)
        {
            worksheet.SheetFormatProperties.CustomHeight = true;
        }
        else
        {
            worksheet.SheetFormatProperties.CustomHeight = null;
        }

        double worksheetColumnWidth = GetColumnWidth(xlWorksheet.ColumnWidth).SaveRound();
        if (xlWorksheet.ColumnWidthChanged)
        {
            worksheet.SheetFormatProperties.DefaultColumnWidth = worksheetColumnWidth;
        }
        else
        {
            worksheet.SheetFormatProperties.DefaultColumnWidth = null;
        }

        if (maxOutlineColumn > 0)
        {
            worksheet.SheetFormatProperties.OutlineLevelColumn = (byte)maxOutlineColumn;
        }
        else
        {
            worksheet.SheetFormatProperties.OutlineLevelColumn = null;
        }

        if (maxOutlineRow > 0)
        {
            worksheet.SheetFormatProperties.OutlineLevelRow = (byte)maxOutlineRow;
        }
        else
        {
            worksheet.SheetFormatProperties.OutlineLevelRow = null;
        }

        #endregion SheetFormatProperties

        #region Columns

        uint worksheetStyleId = context.GetStyleId(xlWorksheet.FormatValue);
        if (
            xlWorksheet.Internals.CellsCollection.IsEmpty
            && xlWorksheet.Internals.ColumnsCollection.Count == 0
            && worksheetStyleId == 0
        )
        {
            worksheet.RemoveAllChildren<Columns>();
        }
        else
        {
            if (!worksheet.Elements<Columns>().Any())
            {
                OpenXmlElement previousElement = cm.GetPreviousElementFor(
                    XLWorksheetContents.Columns
                );
                worksheet.InsertAfter(new Columns(), previousElement);
            }

            Columns columns = worksheet.Elements<Columns>().First();
            cm.SetElement(XLWorksheetContents.Columns, columns);

            Dictionary<uint, Column> sheetColumnsByMin = columns
                .Elements<Column>()
                .ToDictionary(c => c.Min.Value, c => c);
            //Dictionary<UInt32, Column> sheetColumnsByMax = columns.Elements<Column>().ToDictionary(c => c.Max.Value, c => c);

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

            if (minInColumnsCollection > 1)
            {
                UInt32Value min = 1;
                UInt32Value max = (uint)(minInColumnsCollection - 1);

                for (UInt32Value co = min; co <= max; co++)
                {
                    Column column = new()
                    {
                        Min = co,
                        Max = co,
                        Style = worksheetStyleId,
                        Width = worksheetColumnWidth,
                        CustomWidth = true,
                    };

                    UpdateColumn(column, columns, sheetColumnsByMin); //, sheetColumnsByMax);
                }
            }

            for (int co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
            {
                uint styleId;
                double columnWidth;
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
                else
                {
                    styleId = worksheetStyleId;
                    columnWidth = worksheetColumnWidth;
                }

                Column column = new()
                {
                    Min = (uint)co,
                    Max = (uint)co,
                    Style = styleId,
                    Width = columnWidth,
                    CustomWidth = true,
                };

                if (isHidden)
                {
                    column.Hidden = true;
                }

                if (collapsed)
                {
                    column.Collapsed = true;
                }

                if (outlineLevel > 0)
                {
                    column.OutlineLevel = (byte)outlineLevel;
                }

                UpdateColumn(column, columns, sheetColumnsByMin); //, sheetColumnsByMax);
            }

            int collection = maxInColumnsCollection;
            foreach (
                Column col in columns
                    .Elements<Column>()
                    .Where(c => c.Min > (uint)(collection))
                    .OrderBy(c => c.Min.Value)
            )
            {
                col.Style = worksheetStyleId;
                col.Width = worksheetColumnWidth;
                col.CustomWidth = true;

                if ((int)col.Max.Value > maxInColumnsCollection)
                {
                    maxInColumnsCollection = (int)col.Max.Value;
                }
            }

            if (
                maxInColumnsCollection < XlsxSharp.XLHelper.MaxColumnNumber
                && worksheetStyleId != 0
            )
            {
                Column column = new()
                {
                    Min = (uint)(maxInColumnsCollection + 1),
                    Max = (uint)(XlsxSharp.XLHelper.MaxColumnNumber),
                    Style = worksheetStyleId,
                    Width = worksheetColumnWidth,
                    CustomWidth = true,
                };
                columns.AppendChild(column);
            }

            CollapseColumns(columns, sheetColumnsByMin);

            if (!columns.Any())
            {
                worksheet.RemoveAllChildren<Columns>();
                cm.SetElement(XLWorksheetContents.Columns, null);
            }
        }

        #endregion Columns

        #region SheetData

        if (!worksheet.Elements<SheetData>().Any())
        {
            OpenXmlElement previousElement = cm.GetPreviousElementFor(
                XLWorksheetContents.SheetData
            );
            worksheet.InsertAfter(new SheetData(), previousElement);
        }

        SheetData sheetData = worksheet.Elements<SheetData>().First();
        cm.SetElement(XLWorksheetContents.SheetData, sheetData);

        // Sheet data is not updated in the Worksheet DOM here, because it is later being streamed directly to the file
        // without an intermediate DOM representation. This is done to save memory, which is especially problematic
        // for large sheets.

        #endregion SheetData


        #region Tables

        PopulateTablePartReferences((XLTables)xlWorksheet.Tables, worksheet, cm);

        #endregion Tables

        #region Drawings

        if (worksheetPart.DrawingsPart != null)
        {
            XLPictures xlPictures = xlWorksheet.Pictures as Drawings.XLPictures;
            foreach (string removedPicture in xlPictures.Deleted)
            {
                worksheetPart.DrawingsPart.DeletePart(removedPicture);
            }
            xlPictures.Deleted.Clear();
        }

        foreach (XLPicture pic in xlWorksheet.Pictures)
        {
            AddPictureAnchor(worksheetPart, pic, context);
        }

        if (xlWorksheet.Pictures.Count > 0)
        {
            RebaseNonVisualDrawingPropertiesIds(worksheetPart);
        }

        TableParts tableParts = worksheet.Elements<TableParts>().First();
        if (xlWorksheet.Pictures.Count > 0 && !worksheet.OfType<Drawing>().Any())
        {
            Drawing worksheetDrawing = new()
            {
                Id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart),
            };
            worksheetDrawing.AddNamespaceDeclaration(
                "r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            );
            worksheet.InsertBefore(worksheetDrawing, tableParts);
            cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
        }

        // Instead of saving a file with an empty Drawings.xml file, rather remove the .xml file
        bool hasCharts =
            worksheetPart.DrawingsPart is not null && worksheetPart.DrawingsPart.Parts.Any();
        if (
            worksheetPart.DrawingsPart is not null
            && // There is a drawing part for the sheet that could be deleted
            xlWorksheet.LegacyDrawingId is null
            && // and sheet doesn't contain any form controls or comments or other shapes
            xlWorksheet.Pictures.Count == 0
            && // and also no pictures.
            !hasCharts
        ) // and no charts
        {
            string id = worksheetPart.GetIdOfPart(worksheetPart.DrawingsPart);
            worksheet.RemoveChild(worksheet.OfType<Drawing>().FirstOrDefault(p => p.Id == id));
            worksheetPart.DeletePart(worksheetPart.DrawingsPart);
            cm.SetElement(XLWorksheetContents.Drawing, null);
        }

        #endregion Drawings

        #region LegacyDrawing

        // Does worksheet have any comments (stored in legacy VML drawing)
        if (!string.IsNullOrEmpty(xlWorksheet.LegacyDrawingId))
        {
            worksheet.RemoveAllChildren<LegacyDrawing>();
            OpenXmlElement previousElement = cm.GetPreviousElementFor(
                XLWorksheetContents.LegacyDrawing
            );
            worksheet.InsertAfter(
                new LegacyDrawing { Id = xlWorksheet.LegacyDrawingId },
                previousElement
            );

            cm.SetElement(
                XLWorksheetContents.LegacyDrawing,
                worksheet.Elements<LegacyDrawing>().First()
            );
        }
        else
        {
            worksheet.RemoveAllChildren<LegacyDrawing>();
            cm.SetElement(XLWorksheetContents.LegacyDrawing, null);
        }

        #endregion LegacyDrawing

        #region LegacyDrawingHeaderFooter

        //LegacyDrawingHeaderFooter legacyHeaderFooter = worksheetPart.Worksheet.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault();
        //if (legacyHeaderFooter != null)
        //{
        //    worksheetPart.Worksheet.RemoveAllChildren<LegacyDrawingHeaderFooter>();
        //    {
        //            var previousElement = cm.GetPreviousElementFor(XLWSContentManager.XLWSContents.LegacyDrawingHeaderFooter);
        //            worksheetPart.Worksheet.InsertAfter(new LegacyDrawingHeaderFooter { Id = xlWorksheet.LegacyDrawingId },
        //                                                previousElement);
        //    }
        //}

        #endregion LegacyDrawingHeaderFooter

        return worksheet;
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

    private static void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
    {
        uint lastMin = 1;
        int count = sheetColumns.Count;
        KeyValuePair<uint, Column>[] arr = [.. sheetColumns.OrderBy(kp => kp.Key)];
        // sheetColumns[kp.Key + 1]
        //Int32 i = 0;
        //foreach (KeyValuePair<uint, Column> kp in arr
        //    //.Where(kp => !(kp.Key < count && ColumnsAreEqual(kp.Value, )))
        //    )
        for (int i = 0; i < count; i++)
        {
            KeyValuePair<uint, Column> kp = arr[i];
            if (i + 1 != count && ColumnsAreEqual(kp.Value, arr[i + 1].Value))
            {
                continue;
            }

            Column newColumn = (Column)kp.Value.CloneNode(true);
            newColumn.Min = lastMin;
            uint newColumnMax = newColumn.Max.Value;
            List<Column> columnsToRemove =
            [
                .. columns
                    .Elements<Column>()
                    .Where(co => co.Min >= lastMin && co.Max <= newColumnMax)
                    .Select(co => co),
            ];
            columnsToRemove.ForEach(c => columns.RemoveChild(c));

            columns.AppendChild(newColumn);
            lastMin = kp.Key + 1;
            //i++;
        }
    }

    private static double GetColumnWidth(double columnWidth) =>
        Math.Min(255.0, Math.Max(0.0, columnWidth + XLConstants.ColumnWidthOffset));

    private static void UpdateColumn(
        Column column,
        Columns columns,
        Dictionary<uint, Column> sheetColumnsByMin
    )
    {
        if (!sheetColumnsByMin.TryGetValue(column.Min.Value, out Column newColumn))
        {
            newColumn = (Column)column.CloneNode(true);
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(column.Min.Value, newColumn);
        }
        else
        {
            Column existingColumn = sheetColumnsByMin[column.Min.Value];
            newColumn = (Column)existingColumn.CloneNode(true);
            newColumn.Min = column.Min;
            newColumn.Max = column.Max;
            newColumn.Style = column.Style;
            newColumn.Width = column.Width.SaveRound();
            newColumn.CustomWidth = column.CustomWidth;

            if (column.Hidden != null)
            {
                newColumn.Hidden = true;
            }
            else
            {
                newColumn.Hidden = null;
            }

            if (column.Collapsed != null)
            {
                newColumn.Collapsed = true;
            }
            else
            {
                newColumn.Collapsed = null;
            }

            if (column.OutlineLevel != null && column.OutlineLevel > 0)
            {
                newColumn.OutlineLevel = (byte)column.OutlineLevel;
            }
            else
            {
                newColumn.OutlineLevel = null;
            }

            sheetColumnsByMin.Remove(column.Min.Value);
            if (existingColumn.Min + 1 > existingColumn.Max)
            {
                //existingColumn.Min = existingColumn.Min + 1;
                //columns.InsertBefore(existingColumn, newColumn);
                //existingColumn.Remove();
                columns.RemoveChild(existingColumn);
                columns.AppendChild(newColumn);
                sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
            }
            else
            {
                //columns.InsertBefore(existingColumn, newColumn);
                columns.AppendChild(newColumn);
                sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
                existingColumn.Min = existingColumn.Min + 1;
                sheetColumnsByMin.Add(existingColumn.Min.Value, existingColumn);
            }
        }
    }

    private static bool ColumnsAreEqual(Column left, Column right) =>
        (
            (left.Style == null && right.Style == null)
            || (left.Style != null && right.Style != null && left.Style.Value == right.Style.Value)
        )
        && (
            (left.Width == null && right.Width == null)
            || (
                left.Width != null
                && right.Width != null
                && (Math.Abs(left.Width.Value - right.Width.Value) < XlsxSharp.XLHelper.Epsilon)
            )
        )
        && (
            (left.Hidden == null && right.Hidden == null)
            || (
                left.Hidden != null
                && right.Hidden != null
                && left.Hidden.Value == right.Hidden.Value
            )
        )
        && (
            (left.Collapsed == null && right.Collapsed == null)
            || (
                left.Collapsed != null
                && right.Collapsed != null
                && left.Collapsed.Value == right.Collapsed.Value
            )
        )
        && (
            (left.OutlineLevel == null && right.OutlineLevel == null)
            || (
                left.OutlineLevel != null
                && right.OutlineLevel != null
                && left.OutlineLevel.Value == right.OutlineLevel.Value
            )
        );

    // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
    // http://archive.oreilly.com/pub/post/what_is_an_emu.html
    // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
    private static long ConvertToEnglishMetricUnits(int pixels, double resolution) =>
        Convert.ToInt64(914400L * pixels / resolution);

    private static void AddPictureAnchor(
        WorksheetPart worksheetPart,
        Drawings.IXLPicture picture,
        SaveContext context
    )
    {
        XLPicture pic = picture as Drawings.XLPicture;
        DrawingsPart drawingsPart =
            worksheetPart.DrawingsPart
            ?? worksheetPart.AddNewPart<DrawingsPart>(
                context.RelIdGenerator.GetNext(RelType.Workbook)
            );

        if (drawingsPart.WorksheetDrawing == null)
        {
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        }

        Xdr.WorksheetDrawing worksheetDrawing = drawingsPart.WorksheetDrawing;

        // Add namespaces
        if (
            !worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/main")
            )
        )
        {
            worksheetDrawing.AddNamespaceDeclaration(
                "a",
                "http://schemas.openxmlformats.org/drawingml/2006/main"
            );
        }

        if (
            !worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                )
            )
        )
        {
            worksheetDrawing.AddNamespaceDeclaration(
                "r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            );
        }
        /////////

        // Overwrite actual image binary data
        ImagePart imagePart;
        if (drawingsPart.HasPartWithId(pic.RelId))
        {
            imagePart = drawingsPart.GetPartById(pic.RelId) as ImagePart;
        }
        else
        {
            pic.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            imagePart = drawingsPart.AddImagePart(pic.Format.ToOpenXml(), pic.RelId);
        }

        using (MemoryStream stream = new())
        {
            pic.ImageStream.Position = 0;
            pic.ImageStream.CopyTo(stream);
            stream.Seek(0, SeekOrigin.Begin);
            imagePart.FeedData(stream);
        }
        /////////

        // Clear current anchors
        OpenXmlElement existingAnchor = GetAnchorFromImageId(drawingsPart, pic.RelId);

        XLWorkbook wb = pic.Worksheet.Workbook;
        long extentsCx = ConvertToEnglishMetricUnits(pic.Width, wb.DpiX);
        long extentsCy = ConvertToEnglishMetricUnits(pic.Height, wb.DpiY);

        IEnumerable<Xdr.NonVisualDrawingProperties> nvps =
            worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        uint nvpId = nvps.Any()
            ? (UInt32Value)
                worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value)
                + 1
            : 1U;

        Xdr.FromMarker fMark;
        Xdr.ToMarker tMark;
        switch (pic.Placement)
        {
            case Drawings.XLPicturePlacement.FreeFloating:
                Xdr.AbsoluteAnchor absoluteAnchor = new(
                    new Xdr.Position
                    {
                        X = ConvertToEnglishMetricUnits(pic.Left, wb.DpiX),
                        Y = ConvertToEnglishMetricUnits(pic.Top, wb.DpiY),
                    },
                    new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(
                                new PictureLocks { NoChangeAspect = true }
                            )
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print,
                            },
                            new Stretch((FillRectangle)[])
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(absoluteAnchor, existingAnchor);
                break;

            case Drawings.XLPicturePlacement.MoveAndSize:
                XLMarker moveAndSizeFromMarker = pic.Markers[Drawings.XLMarkerPosition.TopLeft];
                if (moveAndSizeFromMarker == null)
                {
                    moveAndSizeFromMarker = new Drawings.XLMarker(picture.Worksheet.Cell("A1"));
                }

                fMark = new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId(
                        (moveAndSizeFromMarker.ColumnNumber - 1).ToInvariantString()
                    ),
                    RowId = new Xdr.RowId(
                        (moveAndSizeFromMarker.RowNumber - 1).ToInvariantString()
                    ),
                    ColumnOffset = new Xdr.ColumnOffset(
                        ConvertToEnglishMetricUnits(moveAndSizeFromMarker.Offset.X, wb.DpiX)
                            .ToInvariantString()
                    ),
                    RowOffset = new Xdr.RowOffset(
                        ConvertToEnglishMetricUnits(moveAndSizeFromMarker.Offset.Y, wb.DpiY)
                            .ToInvariantString()
                    ),
                };

                XLMarker moveAndSizeToMarker = pic.Markers[Drawings.XLMarkerPosition.BottomRight];
                if (moveAndSizeToMarker == null)
                {
                    moveAndSizeToMarker = new Drawings.XLMarker(
                        picture.Worksheet.Cell("A1"),
                        new System.Drawing.Point(picture.Width, picture.Height)
                    );
                }

                tMark = new Xdr.ToMarker
                {
                    ColumnId = new Xdr.ColumnId(
                        (moveAndSizeToMarker.ColumnNumber - 1).ToInvariantString()
                    ),
                    RowId = new Xdr.RowId((moveAndSizeToMarker.RowNumber - 1).ToInvariantString()),
                    ColumnOffset = new Xdr.ColumnOffset(
                        ConvertToEnglishMetricUnits(moveAndSizeToMarker.Offset.X, wb.DpiX)
                            .ToInvariantString()
                    ),
                    RowOffset = new Xdr.RowOffset(
                        ConvertToEnglishMetricUnits(moveAndSizeToMarker.Offset.Y, wb.DpiY)
                            .ToInvariantString()
                    ),
                };

                Xdr.TwoCellAnchor twoCellAnchor = new(
                    fMark,
                    tMark,
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(
                                new PictureLocks { NoChangeAspect = true }
                            )
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print,
                            },
                            new Stretch((FillRectangle)[])
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(twoCellAnchor, existingAnchor);
                break;

            case Drawings.XLPicturePlacement.Move:
                XLMarker moveFromMarker = pic.Markers[Drawings.XLMarkerPosition.TopLeft];
                if (moveFromMarker == null)
                {
                    moveFromMarker = new Drawings.XLMarker(picture.Worksheet.Cell("A1"));
                }

                fMark = new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId(
                        (moveFromMarker.ColumnNumber - 1).ToInvariantString()
                    ),
                    RowId = new Xdr.RowId((moveFromMarker.RowNumber - 1).ToInvariantString()),
                    ColumnOffset = new Xdr.ColumnOffset(
                        ConvertToEnglishMetricUnits(moveFromMarker.Offset.X, wb.DpiX)
                            .ToInvariantString()
                    ),
                    RowOffset = new Xdr.RowOffset(
                        ConvertToEnglishMetricUnits(moveFromMarker.Offset.Y, wb.DpiY)
                            .ToInvariantString()
                    ),
                };

                Xdr.OneCellAnchor oneCellAnchor = new(
                    fMark,
                    new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                    new Xdr.Picture(
                        new Xdr.NonVisualPictureProperties(
                            new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = pic.Name },
                            new Xdr.NonVisualPictureDrawingProperties(
                                new PictureLocks { NoChangeAspect = true }
                            )
                        ),
                        new Xdr.BlipFill(
                            new Blip
                            {
                                Embed = drawingsPart.GetIdOfPart(imagePart),
                                CompressionState = BlipCompressionValues.Print,
                            },
                            new Stretch((FillRectangle)[])
                        ),
                        new Xdr.ShapeProperties(
                            new Transform2D(
                                new Offset { X = 0, Y = 0 },
                                new Extents { Cx = extentsCx, Cy = extentsCy }
                            ),
                            new PresetGeometry { Preset = ShapeTypeValues.Rectangle }
                        )
                    ),
                    new Xdr.ClientData()
                );

                AttachAnchor(oneCellAnchor, existingAnchor);
                break;
        }

        void AttachAnchor(OpenXmlElement pictureAnchor, OpenXmlElement existingAnchor)
        {
            if (existingAnchor is not null)
            {
                worksheetDrawing.ReplaceChild(pictureAnchor, existingAnchor);
            }
            else
            {
                worksheetDrawing.Append(pictureAnchor);
            }
        }
    }

    private static void RebaseNonVisualDrawingPropertiesIds(WorksheetPart worksheetPart)
    {
        Xdr.WorksheetDrawing worksheetDrawing = worksheetPart.DrawingsPart.WorksheetDrawing;

        List<Xdr.NonVisualDrawingProperties> toRebase =
        [
            .. worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>(),
        ];

        toRebase.ForEach(nvdpr => nvdpr.Id = Convert.ToUInt32(toRebase.IndexOf(nvdpr) + 1));
    }

    private static void PopulateTablePartReferences(
        XLTables xlTables,
        Worksheet worksheet,
        XLWorksheetContentManager cm
    )
    {
        XLTable emptyTable = xlTables.FirstOrDefault<XLTable>(t => t.DataRange is null);
        if (emptyTable != null)
        {
            throw new EmptyTableException($"Table '{emptyTable.Name}' should have at least 1 row.");
        }

        TableParts tableParts;
        if (worksheet.Elements<TableParts>().Any())
        {
            tableParts = worksheet.Elements<TableParts>().First();
        }
        else
        {
            OpenXmlElement previousElement = cm.GetPreviousElementFor(
                XLWorksheetContents.TableParts
            );
            tableParts = new TableParts();
            worksheet.InsertAfter(tableParts, previousElement);
        }
        cm.SetElement(XLWorksheetContents.TableParts, tableParts);

        xlTables.Deleted.Clear();
        tableParts.RemoveAllChildren();
        foreach (XLTable xlTable in xlTables.Cast<XLTable>())
        {
            tableParts.AppendChild(new TablePart { Id = xlTable.RelId });
        }

        tableParts.Count = (uint)xlTables.Count<XLTable>();
    }

    /// <summary>
    /// Stream detached worksheet DOM to the worksheet part stream.
    /// Replaces the content of the part.
    /// </summary>
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
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context
    )
    {
        foreach (
            HyperlinkRelationship relationship in worksheetPart.HyperlinkRelationships.ToList()
        )
        {
            worksheetPart.DeleteReferenceRelationship(relationship);
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
                worksheetPart.AddHyperlinkRelationship(hyperlink.ExternalAddress, true, relId);
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

    /// <summary>
    /// The sheet as XML, with its namespaces declared where the DOM declares them.
    /// </summary>
    /// <remarks>
    /// Writing out a detached SDK element hoists every namespace its descendants use onto the
    /// element itself, on top of leaving the declarations where they were. The extra ones come
    /// off again: a declaration stays only if the sheet itself declares it, if it is the sheet's
    /// own namespace, or if one of the sheet's attributes needs it.
    /// </remarks>
    private static XElement ToXml(Worksheet worksheetDom)
    {
        XElement worksheet = SpreadsheetXml.FromSdk(worksheetDom);
        HashSet<string> declared =
        [
            .. worksheetDom.NamespaceDeclarations.Select(declaration => declaration.Value),
            worksheet.Name.NamespaceName,
            .. worksheet.Attributes().Select(attribute => attribute.Name.NamespaceName),
        ];

        foreach (
            XAttribute attribute in worksheet
                .Attributes()
                .Where(attribute => attribute.IsNamespaceDeclaration)
                .ToList()
        )
        {
            if (!declared.Contains(attribute.Value))
            {
                attribute.Remove();
            }
        }

        return worksheet;
    }

    private static void StreamToPart(
        XElement worksheet,
        WorksheetPart worksheetPart,
        XLWorksheet xlWorksheet,
        SaveContext context,
        SaveOptions options
    )
    {
        // Worksheet part might have some data, but creating the stream truncates everything.
        using Stream partStream = worksheetPart.GetStream(FileMode.Create);
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

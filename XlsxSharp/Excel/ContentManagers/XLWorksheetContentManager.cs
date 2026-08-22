#nullable disable

using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ContentManagers;

internal enum XLWorksheetContents
{
    SheetProperties = 1,
    SheetDimension = 2,
    SheetViews = 3,
    SheetFormatProperties = 4,
    Columns = 5,
    SheetData = 6,
    SheetCalculationProperties = 7,
    SheetProtection = 8,
    ProtectedRanges = 9,
    Scenarios = 10,
    AutoFilter = 11,
    SortState = 12,
    DataConsolidate = 13,
    CustomSheetViews = 14,
    MergeCells = 15,
    PhoneticProperties = 16,
    ConditionalFormatting = 17,
    DataValidations = 18,
    Hyperlinks = 19,
    PrintOptions = 20,
    PageMargins = 21,
    PageSetup = 22,
    HeaderFooter = 23,
    RowBreaks = 24,
    ColumnBreaks = 25,
    CustomProperties = 26,
    CellWatches = 27,
    IgnoredErrors = 28,
    SmartTags = 29,
    Drawing = 30,
    LegacyDrawing = 31,
    LegacyDrawingHeaderFooter = 32,
    DrawingHeaderFooter = 33,
    Picture = 34,
    OleObjects = 35,
    Controls = 36,
    AlternateContent = 37,
    WebPublishItems = 38,
    TableParts = 39,
    WorksheetExtensionList = 40,
}

internal class XLWorksheetContentManager : XLBaseContentManager<XLWorksheetContents>
{
    public XLWorksheetContentManager(Worksheet opWorksheet)
    {
        this.contents.Add(
            XLWorksheetContents.SheetProperties,
            opWorksheet.Elements<SheetProperties>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetDimension,
            opWorksheet.Elements<SheetDimension>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetViews,
            opWorksheet.Elements<SheetViews>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetFormatProperties,
            opWorksheet.Elements<SheetFormatProperties>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.Columns,
            opWorksheet.Elements<Columns>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetData,
            opWorksheet.Elements<SheetData>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetCalculationProperties,
            opWorksheet.Elements<SheetCalculationProperties>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SheetProtection,
            opWorksheet.Elements<SheetProtection>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.ProtectedRanges,
            opWorksheet.Elements<ProtectedRanges>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.Scenarios,
            opWorksheet.Elements<Scenarios>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.AutoFilter,
            opWorksheet.Elements<AutoFilter>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.SortState,
            opWorksheet.Elements<SortState>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.DataConsolidate,
            opWorksheet.Elements<DataConsolidate>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.CustomSheetViews,
            opWorksheet.Elements<CustomSheetViews>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.MergeCells,
            opWorksheet.Elements<MergeCells>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.PhoneticProperties,
            opWorksheet.Elements<PhoneticProperties>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.ConditionalFormatting,
            opWorksheet.Elements<ConditionalFormatting>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.DataValidations,
            opWorksheet.Elements<DataValidations>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.Hyperlinks,
            opWorksheet.Elements<Hyperlinks>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.PrintOptions,
            opWorksheet.Elements<PrintOptions>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.PageMargins,
            opWorksheet.Elements<PageMargins>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.PageSetup,
            opWorksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.PageSetup>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.HeaderFooter,
            opWorksheet.Elements<HeaderFooter>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.RowBreaks,
            opWorksheet.Elements<RowBreaks>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.ColumnBreaks,
            opWorksheet.Elements<ColumnBreaks>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.CustomProperties,
            opWorksheet
                .Elements<DocumentFormat.OpenXml.Spreadsheet.CustomProperties>()
                .LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.CellWatches,
            opWorksheet.Elements<CellWatches>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.IgnoredErrors,
            opWorksheet.Elements<IgnoredErrors>().LastOrDefault()
        );
        //contents.Add(XLWSContents.SmartTags, opWorksheet.Elements<SmartTags>().LastOrDefault());
        this.contents.Add(
            XLWorksheetContents.Drawing,
            opWorksheet.Elements<Drawing>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.LegacyDrawing,
            opWorksheet.Elements<LegacyDrawing>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.LegacyDrawingHeaderFooter,
            opWorksheet.Elements<LegacyDrawingHeaderFooter>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.DrawingHeaderFooter,
            opWorksheet.Elements<DrawingHeaderFooter>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.Picture,
            opWorksheet.Elements<Picture>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.OleObjects,
            opWorksheet.Elements<OleObjects>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.Controls,
            opWorksheet.Elements<Controls>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.AlternateContent,
            opWorksheet.Elements<AlternateContent>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.WebPublishItems,
            opWorksheet.Elements<WebPublishItems>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.TableParts,
            opWorksheet.Elements<TableParts>().LastOrDefault()
        );
        this.contents.Add(
            XLWorksheetContents.WorksheetExtensionList,
            opWorksheet.Elements<WorksheetExtensionList>().LastOrDefault()
        );
    }
}

#nullable disable

using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxSharp.Excel.ContentManagers;

internal enum XLSheetViewContents
{
    Pane,
    Selection,
    PivotSelection,
    ExtensionList,
}

internal class XLSheetViewContentManager : XLBaseContentManager<XLSheetViewContents>
{
    public XLSheetViewContentManager(SheetView sheetView)
    {
        this.contents.Add(XLSheetViewContents.Pane, sheetView.Elements<Pane>().LastOrDefault());
        this.contents.Add(
            XLSheetViewContents.Selection,
            sheetView.Elements<Selection>().LastOrDefault()
        );
        this.contents.Add(
            XLSheetViewContents.PivotSelection,
            sheetView.Elements<PivotSelection>().LastOrDefault()
        );
        this.contents.Add(
            XLSheetViewContents.ExtensionList,
            sheetView.Elements<ExtensionList>().LastOrDefault()
        );
    }
}

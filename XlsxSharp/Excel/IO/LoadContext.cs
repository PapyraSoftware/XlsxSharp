using XlsxSharp.Excel.ConditionalFormats;
using XlsxSharp.IO;

namespace XlsxSharp.Excel.IO;

internal class LoadContext
{
    /// <summary>
    /// Conditional formats for pivot tables, loaded from sheets. Key is sheet name, value is the
    /// conditional formats.
    /// </summary>
    private readonly Dictionary<string, List<XLConditionalFormat>> _pivotCfs = new(
        XlsxSharp.XLHelper.SheetComparer
    );

    internal void AddPivotTableCf(string sheetName, XLConditionalFormat conditionalFormat)
    {
        if (!this._pivotCfs.TryGetValue(sheetName, out List<XLConditionalFormat>? list))
        {
            list = [];
            this._pivotCfs[sheetName] = list;
        }

        list.Add(conditionalFormat);
    }

    internal XLConditionalFormat GetPivotCf(string sheetName, int priority)
    {
        if (!this._pivotCfs.TryGetValue(sheetName, out List<XLConditionalFormat>? list))
        {
            throw PivotCfNotFoundException(sheetName, priority);
        }

        XLConditionalFormat? pivotCf = list.SingleOrDefault(x => x.Priority == priority);
        if (pivotCf is null)
        {
            throw PivotCfNotFoundException(sheetName, priority);
        }

        return pivotCf;
    }

    private static Exception PivotCfNotFoundException(string sheetName, int priority) =>
        PartStructureException.ExpectedElementNotFound(
            $"conditional formatting for pivot table in sheet {sheetName} with priority {priority}"
        );
}

namespace XlsxSharp.Excel.Sort;

/// <summary>
/// A comparator of two cell value. It uses semantic of a sort feature in Excel:
/// <list>
///   <item>Order by type is number, text, logical, error, blank.</item>
///   <item>Errors are not sorted.</item>
///   <item>Blanks are always last, both in ascending and descending order.</item>
///   <item>Stable sort.</item>
/// </list>
/// </summary>
internal class XLCellValueSortComparer : IComparer<XLCellValue>
{
    private readonly bool _isAscending;
    private readonly bool _interpretBlankAsString;
    private readonly StringComparer _comparer;

    public XLCellValueSortComparer(IXLSortElement sortElement)
    {
        // Detecting current culture is expensive, when called enough time. Keep pre-calculated comparer.
        this._comparer = sortElement.MatchCase
            ? StringComparer.CurrentCulture
            : StringComparer.CurrentCultureIgnoreCase;
        this._isAscending = sortElement.SortOrder == XLSortOrder.Ascending;
        this._interpretBlankAsString = !sortElement.IgnoreBlanks;
    }

    public int Compare(XLCellValue x, XLCellValue y)
    {
        int xTypeOrder = this.GetTypeOrder(x, this._isAscending);
        int yTypeOrder = this.GetTypeOrder(y, this._isAscending);
        if (xTypeOrder != yTypeOrder)
        {
            return xTypeOrder - yTypeOrder;
        }

        return this._isAscending ? this.CompareAsc(x, y) : -this.CompareAsc(x, y);
    }

    private int GetTypeOrder(XLCellValue x, bool asc)
    {
        // Blank is always last, both for asc and desc.
        if (!this._interpretBlankAsString && x.Type == XLDataType.Blank)
        {
            return 4;
        }

        int ascOrder = x.Type switch
        {
            XLDataType.Number => 0,
            XLDataType.DateTime => 0,
            XLDataType.TimeSpan => 0,
            XLDataType.Text => 1,
            XLDataType.Blank => 1, // If we get here, the blank is interpreted as a text.
            XLDataType.Boolean => 2,
            XLDataType.Error => 3,
            _ => throw new NotSupportedException(),
        };
        return asc ? ascOrder : -ascOrder;
    }

    private int CompareAsc(XLCellValue x, XLCellValue y)
    {
        x = this._interpretBlankAsString && x.Type == XLDataType.Blank ? string.Empty : x;
        y = this._interpretBlankAsString && y.Type == XLDataType.Blank ? string.Empty : y;
        switch (x.Type)
        {
            case XLDataType.Blank:
                return 0; // Blanks are not sorted. That doesn't really affect content, but cells still contain other info, e.g. style.

            case XLDataType.Text:
                return this._comparer.Compare(x.GetText(), y.GetText());

            case XLDataType.Boolean:
                return x.GetBoolean().CompareTo(y.GetBoolean());

            case XLDataType.Error:
                return 0; // Errors are never sorted

            case XLDataType.Number:
            case XLDataType.DateTime:
            case XLDataType.TimeSpan:
                return x.GetUnifiedNumber().CompareTo(y.GetUnifiedNumber());

            default:
                throw new NotSupportedException();
        }
    }
}

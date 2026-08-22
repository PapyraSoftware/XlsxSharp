#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Sort;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLAutoFilter : IXLAutoFilter
{
    /// <summary>
    /// Key is column number.
    /// </summary>
    private readonly Dictionary<int, XLFilterColumn> _columns = new();

    internal IReadOnlyDictionary<int, XLFilterColumn> Columns => this._columns;

    #region IXLAutoFilter Members

    public IEnumerable<IXLRangeRow> HiddenRows => this.Range.Rows(r => r.WorksheetRow().IsHidden);

    public bool IsEnabled { get; set; }

    public IXLRange Range { get; set; }

    public int SortColumn { get; set; }

    public bool Sorted { get; set; }

    public XLSortOrder SortOrder { get; set; }

    public IEnumerable<IXLRangeRow> VisibleRows => this.Range.Rows(r => !r.WorksheetRow().IsHidden);

    IXLAutoFilter IXLAutoFilter.Clear() => this.Clear();

    IXLFilterColumn IXLAutoFilter.Column(string columnLetter) => this.Column(columnLetter);

    IXLFilterColumn IXLAutoFilter.Column(int columnNumber) => this.Column(columnNumber);

    IXLAutoFilter IXLAutoFilter.Sort(
        int columnToSortBy,
        XLSortOrder sortOrder,
        bool matchCase,
        bool ignoreBlanks
    ) => this.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);

    public IXLAutoFilter Reapply()
    {
        // Recalculate shown / hidden rows
        IXLRangeRows rows = this.Range.Rows(2, this.Range.RowCount());
        rows.ForEach(row => row.WorksheetRow().Unhide());

        foreach (XLFilterColumn filterColumn in this._columns.Values)
        {
            filterColumn.Refresh();
        }

        foreach (IXLRangeRow row in rows)
        {
            bool rowMatch = true;

            foreach ((int columnIndex, XLFilterColumn column) in this._columns)
            {
                IXLCell cell = row.Cell(columnIndex);
                bool columnFilterMatch = column.Check(cell);
                rowMatch &= columnFilterMatch;

                if (!rowMatch)
                {
                    break;
                }
            }

            if (!rowMatch)
            {
                row.WorksheetRow().Hide();
            }
        }

        return this;
    }

    #endregion IXLAutoFilter Members

    internal XLFilterColumn Column(string columnLetter)
    {
        int columnNumber = XlsxSharp.XLHelper.GetColumnNumberFromLetter(columnLetter);
        if (columnNumber < 1 || columnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnLetter),
                "Column '" + columnLetter + "' is outside the allowed column range."
            );
        }

        return this.Column(columnNumber);
    }

    internal XLFilterColumn Column(int columnNumber)
    {
        if (columnNumber < 1 || columnNumber > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(
                nameof(columnNumber),
                "Column " + columnNumber + " is outside the allowed column range."
            );
        }

        if (!this._columns.TryGetValue(columnNumber, out XLFilterColumn filterColumn))
        {
            filterColumn = new XLFilterColumn(this, columnNumber);
            this._columns.Add(columnNumber, filterColumn);
        }

        return filterColumn;
    }

    internal XLAutoFilter Clear()
    {
        if (!this.IsEnabled)
        {
            return this;
        }

        this.IsEnabled = false;
        foreach (XLFilterColumn filterColumn in this._columns.Values)
        {
            filterColumn.Clear(false);
        }

        foreach (IXLRangeRow row in this.Range.Rows().Where(r => r.RowNumber() > 1))
        {
            row.WorksheetRow().Unhide();
        }

        return this;
    }

    internal XLAutoFilter Set(IXLRangeBase range)
    {
        IXLTable firstOverlappingTable = range.Worksheet.Tables.FirstOrDefault(t =>
            t.RangeUsed().Intersects(range)
        );
        if (firstOverlappingTable != null)
        {
            throw new InvalidOperationException(
                $"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'"
            );
        }

        this.Range = range.AsRange();
        this.IsEnabled = true;
        return this;
    }

    internal XLAutoFilter Sort(
        int columnToSortBy,
        XLSortOrder sortOrder,
        bool matchCase,
        bool ignoreBlanks
    )
    {
        if (!this.IsEnabled)
        {
            throw new InvalidOperationException("Filter has not been enabled.");
        }

        this.Range.Range(this.Range.FirstCell().CellBelow(), this.Range.LastCell())
            .Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);

        this.Sorted = true;
        this.SortOrder = sortOrder;
        this.SortColumn = columnToSortBy;

        this.Reapply();

        return this;
    }
}

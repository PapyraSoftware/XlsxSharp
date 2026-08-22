using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLColumns : IXLColumns
{
    private readonly List<XLColumn> _columnsCollection = [];

    private readonly XLWorkbook _workbook;
    private readonly XLWorksheet? _worksheet;
    private readonly XLWorksheet? _defaultStyleSheet;

    /// <summary>
    /// This object represents all columns of the worksheet, even non-materialized ones.
    /// </summary>
    [MemberNotNullWhen(true, nameof(_worksheet))]
    private bool AllColumnsOfSheet => this._worksheet is not null;

    private bool IsMaterialized => this._lazyEnumerable == null;

    private IEnumerable<XLColumn>? _lazyEnumerable;

    private IEnumerable<XLColumn> Columns =>
        this._lazyEnumerable ?? this._columnsCollection.AsEnumerable();

    /// <summary>
    /// Create a new instance of <see cref="XLColumns"/>.
    /// </summary>
    /// <param name="workbook">Workbook to which all columns belong.</param>
    /// <param name="worksheet">If worksheet is specified it means that the created instance represents
    /// all columns on a worksheet so changing its width will affect all columns.</param>
    /// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
    /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLColumn"/> to support lazy initialization.</param>
    public XLColumns(
        XLWorkbook workbook,
        XLWorksheet? worksheet,
        XLWorksheet? defaultStyleSheet = null,
        IEnumerable<XLColumn>? lazyEnumerable = null
    )
    {
        this._workbook = workbook;
        this._worksheet = worksheet;
        this._defaultStyleSheet = defaultStyleSheet;
        this._lazyEnumerable = lazyEnumerable;
    }

    #region IXLColumns Members

    public IEnumerator<IXLColumn> GetEnumerator() =>
        this.Columns.Cast<IXLColumn>().OrderBy(r => r.ColumnNumber()).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public double Width
    {
        set
        {
            this.Columns.ForEach(c => c.Width = value);

            if (!this.AllColumnsOfSheet)
            {
                return;
            }

            this._worksheet.ColumnWidth = value;
            this._worksheet.Internals.ColumnsCollection.ForEach(c => c.Value.Width = value);
        }
    }

    public void Delete()
    {
        if (this.AllColumnsOfSheet)
        {
            this._worksheet.Internals.ColumnsCollection.Clear();
            this._worksheet.Internals.CellsCollection.Clear();
        }
        else
        {
            Dictionary<IXLWorksheet, List<int>> toDelete = new();
            foreach (XLColumn c in this.Columns)
            {
                if (!toDelete.TryGetValue(c.Worksheet, out List<int> list))
                {
                    list = [];
                    toDelete.Add(c.Worksheet, list);
                }

                list.Add(c.ColumnNumber());
            }

            foreach (KeyValuePair<IXLWorksheet, List<int>> kp in toDelete)
            {
                foreach (int c in kp.Value.OrderByDescending(c => c))
                {
                    kp.Key.Column(c).Delete();
                }
            }
        }
    }

    public IXLColumns AdjustToContents()
    {
        this.Columns.ForEach(c => c.AdjustToContents());
        return this;
    }

    public IXLColumns AdjustToContents(int startRow)
    {
        this.Columns.ForEach(c => c.AdjustToContents(startRow));
        return this;
    }

    public IXLColumns AdjustToContents(int startRow, int endRow)
    {
        this.Columns.ForEach(c => c.AdjustToContents(startRow, endRow));
        return this;
    }

    public IXLColumns AdjustToContents(double minWidth, double maxWidth)
    {
        this.Columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
        return this;
    }

    public IXLColumns AdjustToContents(int startRow, double minWidth, double maxWidth)
    {
        this.Columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
        return this;
    }

    public IXLColumns AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth)
    {
        this.Columns.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
        return this;
    }

    public void Hide() => this.Columns.ForEach(c => c.Hide());

    public void Unhide() => this.Columns.ForEach(c => c.Unhide());

    public void Group() => this.Group(false);

    public void Group(int outlineLevel) => this.Group(outlineLevel, false);

    public void Ungroup() => this.Ungroup(false);

    public void Group(bool collapse) => this.Columns.ForEach(c => c.Group(collapse));

    public void Group(int outlineLevel, bool collapse) =>
        this.Columns.ForEach(c => c.Group(outlineLevel, collapse));

    public void Ungroup(bool ungroupFromAll) =>
        this.Columns.ForEach(c => c.Ungroup(ungroupFromAll));

    public void Collapse() => this.Columns.ForEach(c => c.Collapse());

    public void Expand() => this.Columns.ForEach(c => c.Expand());

    public IXLCells Cells()
    {
        XLCells cells = new(this._workbook, false, XLCellsUsedOptions.All);
        foreach (XLColumn container in this.Columns)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(this._workbook, true, XLCellsUsedOptions.All);
        foreach (XLColumn container in this.Columns)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(bool includeFormats) =>
        this.CellsUsed(includeFormats ? XLCellsUsedOptions.All : XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._workbook, true, options);
        foreach (XLColumn container in this.Columns)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    /// <summary>
    ///   Adds a vertical page break after this column.
    /// </summary>
    public IXLColumns AddVerticalPageBreaks()
    {
        foreach (XLColumn col in this.Columns)
        {
            col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
        }

        return this;
    }

    public IXLStyle Style
    {
        get => this.Format;
        set => this.Format.SetStyle(value);
    }

    internal XLCellFormat Format
    {
        get
        {
            if (this.AllColumnsOfSheet)
            {
                return XLCellFormat.ForWorksheet(this._worksheet);
            }

            return XLCellFormat.ForColumns(this._workbook, this._defaultStyleSheet, this.Columns);
        }
    }

    #endregion IXLColumns Members

    public void Add(XLColumn column)
    {
        this.Materialize();
        this._columnsCollection.Add(column);
    }

    public void CollapseOnly() => this.Columns.ForEach(c => c.Collapsed = true);

    public IXLColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this.Columns.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Select()
    {
        foreach (IXLColumn range in this)
        {
            range.Select();
        }
    }

    private void Materialize()
    {
        if (this.IsMaterialized)
        {
            return;
        }

        this._columnsCollection.AddRange(this.Columns);
        this._lazyEnumerable = null;
    }
}

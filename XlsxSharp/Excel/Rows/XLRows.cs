using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Rows;

internal class XLRows : IXLRows
{
    private readonly List<XLRow> _rowsCollection = [];
    private readonly XLWorkbook _workbook;
    private readonly XLWorksheet? _worksheet;
    private readonly XLWorksheet? _defaultStyleSheet;

    /// <summary>
    /// This object represents all rows of the worksheet, even non-materialized ones.
    /// </summary>
    [MemberNotNullWhen(true, nameof(_worksheet))]
    private bool AllRowsOfSheet => this._worksheet is not null;

    private bool IsMaterialized => this._lazyEnumerable == null;

    private IEnumerable<XLRow>? _lazyEnumerable;
    private IEnumerable<XLRow> Rows => this._lazyEnumerable ?? this._rowsCollection.AsEnumerable();

    /// <summary>
    /// Create a new instance of <see cref="XLRows"/>.
    /// </summary>
    /// <param name="workbook">Workbook of the rows.</param>
    /// <param name="worksheet">If worksheet is specified it means that the created instance represents
    /// all rows on a worksheet so changing its height will affect all rows.</param>
    /// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
    /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XLRow"/> to support lazy initialization.</param>
    public XLRows(
        XLWorkbook workbook,
        XLWorksheet? worksheet,
        XLWorksheet? defaultStyleSheet = null,
        IEnumerable<XLRow>? lazyEnumerable = null
    )
    {
        this._workbook = workbook;
        this._worksheet = worksheet;
        this._defaultStyleSheet = defaultStyleSheet;
        this._lazyEnumerable = lazyEnumerable;
    }

    #region IXLRows Members

    public IEnumerator<IXLRow> GetEnumerator() =>
        this.Rows.Cast<IXLRow>().OrderBy(r => r.RowNumber()).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public double Height
    {
        set
        {
            this.Rows.ForEach(c => c.Height = value);
            if (!this.AllRowsOfSheet)
            {
                return;
            }

            this._worksheet.RowHeight = value;
            this._worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
        }
    }

    public void Delete()
    {
        if (this.AllRowsOfSheet)
        {
            this._worksheet.Internals.RowsCollection.Clear();
            this._worksheet.Internals.CellsCollection.Clear();
        }
        else
        {
            Dictionary<IXLWorksheet, List<int>> toDelete = new();
            foreach (XLRow r in this.Rows)
            {
                if (!toDelete.TryGetValue(r.Worksheet, out List<int> list))
                {
                    list = [];
                    toDelete.Add(r.Worksheet, list);
                }

                list.Add(r.RowNumber());
            }

            foreach (KeyValuePair<IXLWorksheet, List<int>> kp in toDelete)
            {
                foreach (int r in kp.Value.OrderByDescending(r => r))
                {
                    kp.Key.Row(r).Delete();
                }
            }
        }
    }

    public IXLRows AdjustToContents()
    {
        this.Rows.ForEach(r => r.AdjustToContents());
        return this;
    }

    public IXLRows AdjustToContents(int startColumn)
    {
        this.Rows.ForEach(r => r.AdjustToContents(startColumn));
        return this;
    }

    public IXLRows AdjustToContents(int startColumn, int endColumn)
    {
        this.Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
        return this;
    }

    public IXLRows AdjustToContents(double minHeight, double maxHeight)
    {
        this.Rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
        return this;
    }

    public IXLRows AdjustToContents(int startColumn, double minHeight, double maxHeight)
    {
        this.Rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
        return this;
    }

    public IXLRows AdjustToContents(
        int startColumn,
        int endColumn,
        double minHeight,
        double maxHeight
    )
    {
        this.Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
        return this;
    }

    public void Hide() => this.Rows.ForEach(r => r.Hide());

    public void Unhide() => this.Rows.ForEach(r => r.Unhide());

    public void Group() => this.Group(false);

    public void Group(int outlineLevel) => this.Group(outlineLevel, false);

    public void Ungroup() => this.Ungroup(false);

    public void Group(bool collapse) => this.Rows.ForEach(r => r.Group(collapse));

    public void Group(int outlineLevel, bool collapse) =>
        this.Rows.ForEach(r => r.Group(outlineLevel, collapse));

    public void Ungroup(bool ungroupFromAll) => this.Rows.ForEach(r => r.Ungroup(ungroupFromAll));

    public void Collapse() => this.Rows.ForEach(r => r.Collapse());

    public void Expand() => this.Rows.ForEach(r => r.Expand());

    public IXLCells Cells()
    {
        XLCells cells = new(this._workbook, false, XLCellsUsedOptions.AllContents);
        foreach (XLRow container in this.Rows)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed()
    {
        XLCells cells = new(this._workbook, true, XLCellsUsedOptions.AllContents);
        foreach (XLRow container in this.Rows)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options)
    {
        XLCells cells = new(this._workbook, true, options);
        foreach (XLRow container in this.Rows)
        {
            cells.Add(container.RangeAddress);
        }

        return cells;
    }

    public IXLRows AddHorizontalPageBreaks()
    {
        foreach (XLRow row in this.Rows)
        {
            row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
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
            if (this.AllRowsOfSheet)
            {
                return XLCellFormat.ForWorksheet(this._worksheet);
            }

            return XLCellFormat.ForRows(this._workbook, this._defaultStyleSheet, this.Rows);
        }
    }

    #endregion IXLRows Members

    public void Add(XLRow row)
    {
        this.Materialize();
        this._rowsCollection.Add(row);
    }

    public IXLRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this.Rows.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Select()
    {
        foreach (IXLRow range in this)
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

        this._rowsCollection.AddRange(this.Rows);
        this._lazyEnumerable = null;
    }
}

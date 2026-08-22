using System.Collections.Generic;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Excel;

internal class FormulaSlice : ISlice
{
    private readonly XLWorksheet _sheet;
    private readonly XLCalcEngine _engine;
    private readonly Slice<XLCellFormula?> _formulas = new();

    public FormulaSlice(XLWorksheet sheet)
    {
        this._sheet = sheet;
        this._engine = sheet.Workbook.CalcEngine;
    }

    public bool IsEmpty => this._formulas.IsEmpty;

    public int MaxColumn => this._formulas.MaxColumn;

    public int MaxRow => this._formulas.MaxRow;

    public Dictionary<int, int>.KeyCollection UsedColumns => this._formulas.UsedColumns;

    public IEnumerable<int> UsedRows => this._formulas.UsedRows;

    public void Clear(Area area)
    {
        this._formulas.Clear(area);
    }

    public void DeleteAreaAndShiftLeft(Area areaToDelete)
    {
        this._formulas.DeleteAreaAndShiftLeft(areaToDelete);
    }

    public void DeleteAreaAndShiftUp(Area areaToDelete)
    {
        this._formulas.DeleteAreaAndShiftUp(areaToDelete);
    }

    public IEnumerator<Point> GetEnumerator(Area area, bool reverse = false)
    {
        return this._formulas.GetEnumerator(area, reverse);
    }

    public void InsertAreaAndShiftDown(Area areaToInsert)
    {
        this._formulas.InsertAreaAndShiftDown(areaToInsert);
    }

    public void InsertAreaAndShiftRight(Area areaToInsert)
    {
        this._formulas.InsertAreaAndShiftRight(areaToInsert);
    }

    public bool IsUsed(Point address)
    {
        return this._formulas.IsUsed(address);
    }

    public void Swap(Point sp1, Point sp2)
    {
        XLCellFormula? value1 = this._formulas[sp1];
        XLCellFormula? value2 = this._formulas[sp2];

        value1 = value1?.GetMovedTo(sp1, sp2);
        value2 = value2?.GetMovedTo(sp2, sp1);

        this.Set(sp1, value2);
        this.Set(sp2, value1);
    }

    internal XLCellFormula? Get(Point point)
    {
        return this._formulas[point];
    }

    internal void Set(Point point, XLCellFormula? formula)
    {
        // Can't ref, because it is an alias for a memory and thus wouldn't hold old formula.
        XLCellFormula? original = this._formulas[point];
        if (ReferenceEquals(original, formula))
        {
            return;
        }

        this._formulas.Set(point, formula);

        // Remove first, so calc chain doesn't choke on two formulas
        // in one cell when changing a formula of a cell.
        SheetPoint bookPoint = new(this._sheet.Name, point);
        if (original is not null)
        {
            this._engine.RemoveFormula(bookPoint, original);
        }

        if (formula is not null)
        {
            this._engine.AddNormalFormula(
                bookPoint,
                this._sheet.Name,
                formula,
                this._sheet.Workbook
            );
        }
    }

    /// <summary>
    /// Set all cells in a <paramref name="range"/> to the array formula.
    /// </summary>
    /// <remarks>
    /// This method doesn't check that formula doesn't damage other array formulas.
    /// </remarks>
    internal void SetArray(Area range, XLCellFormula? arrayFormula)
    {
        for (int row = range.TopRow; row <= range.BottomRow; ++row)
        {
            for (int col = range.LeftColumn; col <= range.RightColumn; ++col)
            {
                Point point = new(row, col);
                XLCellFormula? original = this._formulas[point];

                this._formulas.Set(point, arrayFormula);

                // The formula removal removes formula from dependency tree
                // (number of cells formula affects doesn't matter) and also
                // removes point from the calc chain. Therefore, it works for
                // array and normal formulas.
                SheetPoint bookPoint = new(this._sheet.Name, point);
                if (original is not null)
                {
                    this._engine.RemoveFormula(bookPoint, original);
                }
            }
        }

        if (arrayFormula is not null)
        {
            this._engine.AddArrayFormula(range, arrayFormula, this._sheet);
        }
    }

    internal Slice<XLCellFormula>.Enumerator GetForwardEnumerator(Area range)
    {
        return new Slice<XLCellFormula>.Enumerator(this._formulas!, range);
    }

    /// <summary>
    /// Mark all formulas in a range as dirty.
    /// </summary>
    internal void MarkDirty(Area range)
    {
        using Slice<XLCellFormula>.Enumerator enumerator = this.GetForwardEnumerator(range);
        while (enumerator.MoveNext())
        {
            enumerator.Current.IsDirty = true;
        }
    }
}

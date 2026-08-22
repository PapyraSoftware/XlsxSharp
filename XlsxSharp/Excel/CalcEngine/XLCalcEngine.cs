using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XlsxSharp.Excel.CalcEngine.Exceptions;
using XlsxSharp.Excel.CalcEngine.Functions;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// CalcEngine parses strings and returns Expression objects that can
/// be evaluated.
/// </summary>
/// <remarks>
/// <para>This class has three extensibility points:</para>
/// <para>Use the <b>RegisterFunction</b> method to define custom functions.</para>
/// </remarks>
internal class XLCalcEngine : ISheetListener, IWorkbookListener
{
    private readonly CultureInfo _culture;
    private readonly FormulaParser _parser;
    private readonly CalculationVisitor _visitor;
    private DependencyTree? _dependencyTree;
    private XLCalculationChain? _chain;

    public XLCalcEngine(CultureInfo culture)
    {
        this._culture = culture;
        FunctionRegistry funcRegistry = GetFunctionTable();
        this._parser = new FormulaParser(funcRegistry);
        this._visitor = new CalculationVisitor(funcRegistry);
        this._dependencyTree = null;
        this._chain = null;
    }

    /// <summary>
    /// Parses a string into an <see cref="Formula"/>.
    /// </summary>
    /// <param name="expression">String to parse.</param>
    /// <returns>An formula that can be evaluated.</returns>
    public Formula Parse(string expression)
    {
        return this._parser.GetAst(expression, isA1: true);
    }

    /// <summary>
    /// Add an array formula to the calc engine to manage dirty tracking and evaluation.
    /// </summary>
    internal void AddArrayFormula(Area range, XLCellFormula arrayFormula, XLWorksheet sheet)
    {
        if (this._chain is not null && this._dependencyTree is not null)
        {
            SheetArea area = new(sheet.Name, range);
            this._dependencyTree.AddFormula(area, arrayFormula, sheet.Workbook);
            this._chain.AppendArea(area);
        }
    }

    /// <summary>
    /// Add a formula to the calc engine to manage dirty tracking and evaluation.
    /// </summary>
    internal void AddNormalFormula(
        SheetPoint point,
        string sheetName,
        XLCellFormula formula,
        XLWorkbook workbook
    )
    {
        if (this._chain is not null && this._dependencyTree is not null)
        {
            SheetArea pointArea = new(sheetName, new Area(point.Point, point.Point));
            this._dependencyTree.AddFormula(pointArea, formula, workbook);
            this._chain.AddLast(point);
        }
    }

    /// <summary>
    /// Remove formula from dependency tree (=precedents won't mark
    /// it as dirty) and remove <paramref name="point"/> from the chain.
    /// Note that even if formula is used by many cells (e.g. array formula),
    /// it is fully removed from dependency tree, but each cells referencing
    /// the formula must be removed individually from calc chain.
    /// </summary>
    internal void RemoveFormula(SheetPoint point, XLCellFormula formula)
    {
        if (this._chain is not null && this._dependencyTree is not null)
        {
            this._dependencyTree.RemoveFormula(formula);
            this._chain.Remove(point);
        }
    }

    internal void OnAddedSheet(XLWorksheet sheet)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    internal void OnDeletingSheet(XLWorksheet sheet)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnInsertAreaAndShiftDown(XLWorksheet sheet, Area area)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnInsertAreaAndShiftRight(XLWorksheet sheet, Area area)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnDeleteAreaAndShiftLeft(XLWorksheet sheet, Area deletedArea)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    public void OnDeleteAreaAndShiftUp(XLWorksheet sheet, Area deletedArea)
    {
        this.Purge(sheet.Workbook.WorksheetsInternal);
    }

    private void Purge(XLWorksheets sheets)
    {
        this._dependencyTree = null;
        this._chain = null;

        // Mark everything as dirty, because there can be stale values
        foreach (XLWorksheet sheet in sheets)
        {
            sheet.Internals.CellsCollection.FormulaSlice.MarkDirty(Area.Full);
        }
    }

    internal void MarkDirty(XLWorksheet sheet, Point point)
    {
        this.MarkDirty(sheet, new Area(point, point));
    }

    internal void MarkDirty(XLWorksheet sheet, Area area)
    {
        if (this._dependencyTree is not null)
        {
            SheetArea bookArea = new(sheet.Name, area);
            this._dependencyTree.MarkDirty(bookArea);
        }
    }

    /// <summary>
    /// Recalculate a workbook or a sheet.
    /// </summary>
    internal void Recalculate(XLWorkbook wb, string? recalculateSheetName)
    {
        // Lazy, so initialize chain from wb, if it is empty
        if (this._chain is null || this._dependencyTree is null)
        {
            this._chain = XLCalculationChain.CreateFrom(wb);
            this._dependencyTree = DependencyTree.CreateFrom(wb);
        }

        Dictionary<
            string,
            (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)
        > sheetMap = wb.WorksheetsInternal.ToDictionary<
            XLWorksheet,
            string,
            (XLWorksheet Sheet, ValueSlice ValueSlice, FormulaSlice FormulaSlice)
        >(
            sheet => sheet.Name,
            sheet =>
                (
                    sheet,
                    sheet.Internals.CellsCollection.ValueSlice,
                    sheet.Internals.CellsCollection.FormulaSlice
                ),
            XlsxSharp.XLHelper.SheetComparer
        );

        // Each outer loop moves chain one cell ahead.
        while (this._chain.MoveAhead())
        {
            // Inner loop that pushes supporting formulas ahead of current.
            // It ends when a cell has been calculated and thus chain can move ahead.
            while (true)
            {
                SheetPoint current = this._chain.Current;
                string sheetName = current.SheetName;

                // Skip dirty cells from sheets that are not being recalculated
                if (
                    recalculateSheetName is not null
                    && !XlsxSharp.XLHelper.SheetComparer.Equals(sheetName, recalculateSheetName)
                )
                {
                    // Even though cell is dirty, it's in the ignored sheet and
                    // thus chain can move ahead.
                    break;
                }

                if (
                    !sheetMap.TryGetValue(
                        sheetName,
                        out (
                            XLWorksheet Sheet,
                            ValueSlice ValueSlice,
                            FormulaSlice FormulaSlice
                        ) sheetInfo
                    )
                )
                {
                    throw new InvalidOperationException(
                        $"Sheet of a calc chain cell {current} doesn't exist in the workbook."
                    );
                }

                if (this._chain.IsCurrentInCycle)
                {
                    throw new InvalidOperationException(
                        $"Formula in a cell {current} is part of a cycle."
                    );
                }

                XLCellFormula? cellFormula = sheetInfo.FormulaSlice.Get(current.Point);
                if (cellFormula is null)
                {
                    throw new InvalidOperationException(
                        $"Calculation chain contains a {current}, but the cell doesn't contain a formula."
                    );
                }

                if (!cellFormula.IsDirty)
                {
                    break;
                }

                try
                {
                    this.ApplyFormula(
                        cellFormula,
                        current.Point,
                        sheetInfo.Sheet,
                        sheetInfo.ValueSlice,
                        recalculateSheetName
                    );
                    cellFormula.IsDirty = false;

                    // Break out of the inner loop, a dirty cell has been
                    // calculated and thus chain can move ahead.
                    break;
                }
                catch (GettingDataException ex)
                {
                    this._chain.MoveToCurrent(ex.Point);
                }
            }
        }

        // Super important to clean up the chain for next recalculation.
        // Chain contains shared data and not cleaning it would cause hard
        // to diagnose issues.
        this._chain.Reset();
    }

    private void ApplyFormula(
        XLCellFormula formula,
        Point appliedPoint,
        XLWorksheet sheet,
        ValueSlice valueSlice,
        string? recalculateSheetName
    )
    {
        string formulaText = formula.A1;
        if (formula.Type == FormulaType.Normal)
        {
            ScalarValue single = this.EvaluateFormula(
                formulaText,
                sheet.Workbook,
                sheet,
                new XLAddress(sheet, appliedPoint.Row, appliedPoint.Column, true, true),
                recalculateSheetName: recalculateSheetName
            );
            valueSlice.SetCellValue(appliedPoint, single.ToCellValue());
        }
        else if (formula.Type == FormulaType.Array)
        {
            // The point can be any point in an array, so we can't use it.
            Area range = formula.Range;
            Point leftTopCorner = range.FirstPoint;
            XLCell masterCell = sheet.Cell(leftTopCorner.Row, leftTopCorner.Column);
            Array array = this.EvaluateArrayFormula(formulaText, masterCell, recalculateSheetName);

            // The array from formula can be smaller or larger than the
            // range of cells it should fit into. Broadcast it to the size.
            Array result = array.Broadcast(range.Height, range.Width);

            // Copy value to the value slice
            for (int rowIdx = 0; rowIdx < result.Height; ++rowIdx)
            {
                for (int colIdx = 0; colIdx < result.Width; ++colIdx)
                {
                    ScalarValue cellValue = result[rowIdx, colIdx];
                    int row = range.FirstPoint.Row + rowIdx;
                    int column = range.FirstPoint.Column + colIdx;
                    valueSlice.SetCellValue(new Point(row, column), cellValue.ToCellValue());
                }
            }
        }
        else
        {
            throw new NotImplementedException(
                $"Evaluation of formula type '{formula.Type}' is not supported."
            );
        }
    }

    /// <summary>
    /// Evaluates a normal formula.
    /// </summary>
    /// <param name="expression">Expression to evaluate.</param>
    /// <param name="wb">Workbook where is formula being evaluated.</param>
    /// <param name="ws">Worksheet where is formula being evaluated.</param>
    /// <param name="address">Address of formula.</param>
    /// <param name="recursive">Should the data necessary for this formula (not deeper ones)
    /// be calculated recursively? Used only for non-cell calculations.</param>
    /// <param name="recalculateSheetName">
    /// If set, calculation  will allow dirty reads from other sheets than the passed one.
    /// </param>
    /// <returns>The value of the expression.</returns>
    /// <remarks>
    /// If you are going to evaluate the same expression several times,
    /// it is more efficient to parse it only once using the <see cref="Parse"/>
    /// method and then using the Expression.Evaluate method to evaluate
    /// the parsed expression.
    /// </remarks>
    internal ScalarValue EvaluateFormula(
        string expression,
        XLWorkbook? wb = null,
        XLWorksheet? ws = null,
        IXLAddress? address = null,
        bool recursive = false,
        string? recalculateSheetName = null
    )
    {
        CalcContext ctx = new(this, this._culture, wb, ws, address, recursive)
        {
            RecalculateSheetName = recalculateSheetName,
        };
        AnyValue result = this.EvaluateFormula(expression, ctx);
        if (CalcContext.UseImplicitIntersection)
        {
            result = result.Match(
                () => AnyValue.Blank,
                logical => logical,
                number => number,
                text => text,
                error => error,
                array => array[0, 0].ToAnyValue(),
                reference => reference
            );
        }

        return ToCellContentValue(result, ctx);
    }

    private Array EvaluateArrayFormula(
        string expression,
        XLCell masterCell,
        string? recalculateSheetName
    )
    {
        CalcContext ctx = new(this, this._culture, masterCell)
        {
            IsArrayCalculation = true,
            RecalculateSheetName = recalculateSheetName,
        };
        AnyValue result = this.EvaluateFormula(expression, ctx);
        if (result.TryPickSingleOrMultiValue(out ScalarValue single, out Array multi, ctx))
        {
            return new ScalarArray(single, 1, 1);
        }

        return multi;
    }

    internal AnyValue EvaluateName(string nameFormula, XLWorksheet ws)
    {
        CalcContext ctx = new(this, this._culture, ws.Workbook, ws, null);
        return this.EvaluateFormula(nameFormula, ctx);
    }

    private AnyValue EvaluateFormula(string expression, CalcContext ctx)
    {
        Formula formula = this.Parse(expression);
        AnyValue result = formula.AstRoot.Accept(ctx, this._visitor);
        return result;
    }

    // build/get static keyword table
    private static FunctionRegistry GetFunctionTable()
    {
        FunctionRegistry fr = new();

        // register built-in functions (and constants)
        Engineering.Register(fr);
        Information.Register(fr);
        Logical.Register(fr);
        Lookup.Register(fr);
        MathTrig.Register(fr);
        Text.Register(fr);
        Statistical.Register(fr);
        DateAndTime.Register(fr);
        Financial.Register(fr);

        return fr;
    }

    /// <summary>
    /// Convert any kind of formula value to value returned as a content of a cell.
    /// <list type="bullet">
    ///    <item><c>bool</c> - represents a logical value.</item>
    ///    <item><c>double</c> - represents a number and also date/time as serial date-time.</item>
    ///    <item><c>string</c> - represents a text value.</item>
    ///    <item><see cref="XLError" /> - represents a formula calculation error.</item>
    /// </list>
    /// </summary>
    private static ScalarValue ToCellContentValue(AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
        {
            return scalar;
        }

        if (collection.TryPickT0(out Array? array, out Reference? reference))
        {
            return array![0, 0];
        }

        if (reference!.TryGetSingleCellValue(out ScalarValue cellValue, ctx))
        {
            return cellValue;
        }

        OneOf<Reference, XLError> intersected = reference.ImplicitIntersection(ctx.FormulaAddress);
        if (!intersected.TryPickT0(out Reference? singleCellReference, out XLError error))
        {
            return error;
        }

        if (!singleCellReference!.TryGetSingleCellValue(out ScalarValue singleCellValue, ctx))
        {
            throw new InvalidOperationException(
                "Got multi cell reference instead of single cell reference."
            );
        }

        return singleCellValue;
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        if (this._dependencyTree is not null)
        {
            this._dependencyTree.RenameSheet(oldSheetName, newSheetName);
        }
    }
}

internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);

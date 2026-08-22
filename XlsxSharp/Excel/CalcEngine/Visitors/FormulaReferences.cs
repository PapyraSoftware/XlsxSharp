using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Parser;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel.CalcEngine.Visitors;

/// <summary>
/// A collection of all references in the book (not others) found in a formula.
/// Created by <see cref="CollectRefsFactory"/>.
/// </summary>
internal class FormulaReferences
{
    private readonly string _formula;

    private FormulaReferences(string formula) => this._formula = formula;

    /// <summary>
    /// Is there a <c>#REF!</c> anywhere in the formula?
    /// </summary>
    internal bool ContainsRefError { get; set; }

    /// <summary>
    /// Areas without a sheet found in the formula.
    /// </summary>
    internal HashSet<XLReference> References { get; } = [];

    /// <summary>
    /// Areas with a sheet found in the formula.
    /// </summary>
    internal HashSet<XLSheetReference> SheetReferences { get; } = [];

    internal HashSet<(string Table, string Column, string Symbol)> StructuredReferences { get; } =
    [];

    internal static FormulaReferences ForFormula(string formula)
    {
        FormulaReferences references = new(formula);
        FormulaParser<object?, object?, FormulaReferences>.CellFormulaA1(
            formula,
            references,
            CollectRefsFactory.Instance
        );
        return references;
    }

    internal bool ContainsSheet(string worksheetName) =>
        this.SheetReferences.Any(x =>
            XlsxSharp.XLHelper.SheetComparer.Equals(x.Sheet, worksheetName)
        );

    internal XLRanges GetExternalRanges(XLWorkbook workbook, Point anchor)
    {
        XLRanges list = new(workbook);
        foreach (XLSheetReference reference in this.SheetReferences)
        {
            if (workbook.TryGetWorksheet(reference.Sheet, out XLWorksheet sheet))
            {
                XLRangeAddress rangeAddress = reference.Reference.ToRangeAddress(sheet, anchor);
                list.Add(sheet.Range(rangeAddress));
            }
        }

        foreach ((string tableName, string column, string _) in this.StructuredReferences)
        {
            if (workbook.TryGetTable(tableName, out XLTable table))
            {
                list.Add(table.DataRange.Column(column));
            }
        }

        return list;
    }

    /// <summary>
    /// Factory to get all references (cells, tables, names) in local workbook.
    /// </summary>
    private class CollectRefsFactory : CollectVisitor<FormulaReferences>
    {
        public static readonly CollectRefsFactory Instance = new();

        public override object? ErrorNode(
            FormulaReferences context,
            SymbolRange range,
            ReadOnlySpan<char> error
        )
        {
            context.ContainsRefError = true;
            return base.ErrorNode(context, range, error);
        }

        public override object? Reference(
            FormulaReferences context,
            SymbolRange range,
            ReferenceArea reference
        )
        {
            context.References.Add(new XLReference(reference));
            return base.Reference(context, range, reference);
        }

        public override object? SheetReference(
            FormulaReferences context,
            SymbolRange range,
            string sheet,
            ReferenceArea reference
        )
        {
            context.SheetReferences.Add(new XLSheetReference(sheet, new XLReference(reference)));
            return base.SheetReference(context, range, sheet, reference);
        }

        public override object? StructureReference(
            FormulaReferences context,
            SymbolRange range,
            string table,
            StructuredReferenceArea area,
            string? firstColumn,
            string? lastColumn
        )
        {
            // TODO: Temporary placeholder, extract range detection from CalculationVisitor
            if (firstColumn is not null)
            {
                context.StructuredReferences.Add(
                    (
                        table,
                        firstColumn,
                        context._formula.Substring(range.Start, range.End - range.Start)
                    )
                );
            }

            return base.StructureReference(context, range, table, area, firstColumn, lastColumn);
        }
    }
}

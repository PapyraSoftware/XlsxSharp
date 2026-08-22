using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ClosedXML.Parser;
using XlsxSharp.Excel.CalcEngine.Visitors;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

[DebuggerDisplay("{_name}:{_formula}")]
internal class XLDefinedName : IXLDefinedName, IWorkbookListener
{
    private readonly XLDefinedNames _container;
    private string _name;
    private string _formula = null!;
    private FormulaReferences _references = null!;

    internal XLDefinedName(
        XLDefinedNames container,
        string name,
        bool validateName,
        string formula,
        string? comment
    )
    {
        // Excel accepts invalid names per grammar (e.g. `[Foo]Bar`) as a valid name and they can
        // encountered in existing workbooks. We shouldn't throw exception on load.
        if (validateName)
        {
            if (!XlsxSharp.XLHelper.ValidateName("named range", name, out string error))
            {
                throw new ArgumentException(error, nameof(name));
            }
        }

        this._container = container;
        this._name = name;
        this.RefersTo = formula;
        this.Visible = true;
        this.Comment = comment;
    }

    public bool IsValid => !this._references.ContainsRefError;

    public string Name
    {
        get => this._name;
        set
        {
            if (XlsxSharp.XLHelper.NameComparer.Equals(this._name, value))
            {
                return;
            }

            if (!XlsxSharp.XLHelper.ValidateName("named range", value, out string error))
            {
                throw new ArgumentException(error, nameof(value));
            }

            if (this._container.Contains(value))
            {
                throw new InvalidOperationException($"There is already a name '{value}'.");
            }

            this._container.Delete(this._name);
            this._name = value;
            this._container.Add(this._name, this);
        }
    }

    public IXLRanges Ranges =>
        this._references.GetExternalRanges(this._container.Workbook, new Point(1, 1));

    public string? Comment { get; set; }

    public bool Visible { get; set; }

    public XLNamedRangeScope Scope => this._container.Scope;

    public string RefersTo
    {
        get => this._formula;
        set
        {
            ArgumentNullException.ThrowIfNull(value);

            string formula = value.TrimFormulaEqual();
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new ArgumentException("Formula can't be empty.");
            }

            FormulaReferences references = FormulaReferences.ForFormula(formula);
            if (references.References.Any())
            {
                // `[MS-XLSX] 2.2.2.5: The formula MUST NOT use the local-cell-reference production
                // rule.` Excel will refuse to load a workbook with such a defined name (e.g. `A1`).
                // In theory, defined name should support bang references as a replacement for local
                // references, but ClosedParser doesn't support it yet.
                throw new ArgumentException(
                    $"Formula '{formula}' contains references without a sheet."
                );
            }

            this._references = references;
            this._formula = formula;
        }
    }

    IXLDefinedName IXLDefinedName.CopyTo(IXLWorksheet targetSheet) =>
        this.CopyTo((XLWorksheet)targetSheet);

    void IXLDefinedName.Delete() => this._container.Delete(this.Name);

    /// <summary>
    /// Get sheet references found in the formula in A1. Doesn't return tables or name references,
    /// only what has col/row coordinates.
    /// </summary>
    internal IReadOnlyList<string> SheetReferencesList =>
        this._references.SheetReferences.Select(x => x.GetA1()).ToList();

    internal XLDefinedName CopyTo(XLWorksheet targetSheet)
    {
        XLWorksheet? sheet = this._container.Worksheet;
        if (targetSheet == sheet)
        {
            throw new InvalidOperationException(
                "Cannot copy named range to the worksheet it already belongs to."
            );
        }

        if (sheet is null)
        {
            throw new InvalidOperationException("Cannot copy workbook scoped defined name.");
        }

        Dictionary<Area, XLTable> targetTables = targetSheet.Tables.ToDictionary<XLTable, Area>(x =>
            x.SheetRange
        );
        Dictionary<string, string> tableRenames = new();
        foreach (XLTable table in sheet.Tables)
        {
            if (targetTables.TryGetValue(table.SheetRange, out XLTable? targetTable))
            {
                tableRenames.Add(table.Name, targetTable.Name);
            }
        }

        string copiedFormula = FormulaConverter.ModifyA1(
            this._formula,
            sheet.Name,
            1,
            1,
            new RenameRefModVisitor
            {
                Sheets = new Dictionary<string, string?> { { sheet.Name, targetSheet.Name } },
                Tables = tableRenames,
            }
        );
        XLDefinedName copiedName = new(
            targetSheet.DefinedNames,
            this.Name,
            false,
            copiedFormula,
            this.Comment
        );
        return targetSheet.DefinedNames.Add(this.Name, copiedName);
    }

    public IXLDefinedName SetRefersTo(IXLRangeBase range) => this.SetRefersTo(RangeToFixed(range));

    public IXLDefinedName SetRefersTo(IXLRanges ranges)
    {
        string unionFormula = string.Join(",", ranges.Select(RangeToFixed));
        return this.SetRefersTo(unionFormula);
    }

    public IXLDefinedName SetRefersTo(string formula)
    {
        this.RefersTo = formula;
        return this;
    }

    public override string ToString() => this._formula;

    internal void Add(string rangeAddress)
    {
        string[] byExclamation = rangeAddress.Split('!');
        string wsName = byExclamation[0].Replace("'", "");
        string rng = byExclamation[1];
        IXLRange rangeToAdd = this
            ._container.Workbook.WorksheetsInternal.Worksheet(wsName)
            .Range(rng);

        XLRanges ranges = new(this._container.Workbook) { rangeToAdd };
        this.RefersTo =
            this._formula + "," + string.Join(",", ranges.Select<XLRange, string>(RangeToFixed));
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName) =>
        this.RenameFormulaSheet(oldSheetName, newSheetName);

    internal void OnWorksheetDeleted(string worksheetName) =>
        this.RenameFormulaSheet(worksheetName, null);

    private void RenameFormulaSheet(string oldSheetName, string? newSheetName)
    {
        if (!this._references.ContainsSheet(oldSheetName))
        {
            return;
        }

        string modified = FormulaConverter.ModifyA1(
            this._formula,
            newSheetName ?? string.Empty,
            1,
            1,
            new RenameRefModVisitor
            {
                Sheets = new Dictionary<string, string?> { { oldSheetName, newSheetName } },
            }
        );

        this.RefersTo = modified;
    }

    private static string RangeToFixed(IXLRangeBase range) =>
        range.RangeAddress.ToStringFixed(XLReferenceStyle.A1, true);
}

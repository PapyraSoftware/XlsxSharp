using System;
using System.Collections.Generic;
using System.Linq;
using RBush;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// <para>
/// A dependency tree structure to hold all formulas of the workbook and reference
/// objects they depend on. The key feature of dependency tree is to propagate
/// dirty flag across formulas.
/// </para>
/// <para>
/// When a data in a cell changes, all formulas that depend on it should be marked
/// as dirty, but it is hard to find which cells are affected - that is what
/// dependency tree does.
/// </para>
/// <para>
/// Dependency tree must be updated, when structure of a workbook is updated:
/// <list type="bullet">
///   <item>Sheet is added, renamed or deleted.</item>
///   <item>Name is added or deleted.</item>
///   <item>Table is resized, renamed, added or deleted.</item>
/// </list>
/// Any such action changes what cells formula depends on and
/// the formula dependencies must be updated.
/// </para>
/// </summary>
internal class DependencyTree
{
    /// <summary>
    /// The source of the truth, a storage of formula dependencies. The dependency tree is
    /// constructed from this collection.
    /// </summary>
    private readonly Dictionary<XLCellFormula, FormulaDependencies> _dependencies = new();

    /// <summary>
    /// Visitor to extract precedents of formulas.
    /// </summary>
    private readonly DependenciesVisitor _visitor;

    /// <summary>
    /// A dependency tree for each sheet (key is sheet name).
    /// </summary>
    private readonly Dictionary<string, SheetDependencyTree> _sheetTrees = new(
        XlsxSharp.XLHelper.SheetComparer
    );

    public DependencyTree() => this._visitor = new DependenciesVisitor();

    internal bool IsEmpty =>
        this._sheetTrees.All(sheetTree => sheetTree.Value.IsEmpty) && this._dependencies.Count == 0;

    internal static DependencyTree CreateFrom(XLWorkbook workbook)
    {
        DependencyTree tree = new();

        // Add tree before adding formulas, because formula can reference any sheet.
        foreach (XLWorksheet sheet in workbook.WorksheetsInternal)
        {
            tree.AddSheetTree(sheet);
        }

        foreach (XLWorksheet sheet in workbook.WorksheetsInternal)
        {
            using Slice<XLCellFormula>.Enumerator enumerator =
                sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
            {
                XLCellFormula formula = enumerator.Current;
                Point point = enumerator.Point;
                if (formula.Type == FormulaType.Normal)
                {
                    SheetArea bookArea = new(sheet.Name, new Area(point, point));
                    tree.AddFormula(bookArea, formula, workbook);
                }
                else if (formula.Type == FormulaType.Array)
                {
                    // Ignore all non-master cells
                    bool isMasterCell = formula.Range.FirstPoint == point;
                    if (isMasterCell)
                    {
                        SheetArea bookArea = new(sheet.Name, formula.Range);
                        tree.AddFormula(bookArea, formula, workbook);
                    }
                }
                else
                {
                    // TODO: Implement other formulas. Don't throw on data table or shared formulas.
                }
            }
        }

        return tree;
    }

    /// <summary>
    /// Add a formula to the dependency tree.
    /// </summary>
    /// <param name="formulaArea">Area of a formula, for normal cells 1x1, for array can be larger.</param>
    /// <param name="formula">The cell formula.</param>
    /// <param name="workbook">Workbook that is used to find precedents (names ect.).</param>
    /// <returns>Added cell formula dependencies.</returns>
    /// <exception cref="ArgumentException">Formula already is in the tree.</exception>
    internal FormulaDependencies AddFormula(
        SheetArea formulaArea,
        XLCellFormula formula,
        XLWorkbook workbook
    )
    {
        FormulaDependencies precedents = this.GetFormulaPrecedents(formulaArea, formula, workbook);

        this._dependencies.Add(formula, precedents);

        foreach (SheetArea precedentArea in precedents.Areas)
        {
            // Add dependency to its sheet dependency tree. The formula might contain
            // a dependency for a sheet that doesn't exist in a workbook. Such dependencies
            // are ignored, until sheet is added.
            if (
                this._sheetTrees.TryGetValue(precedentArea.Name, out SheetDependencyTree? sheetTree)
            )
            {
                // Dependent worksheet exists
                Dependent dependent = new(formulaArea, formula);
                sheetTree.AddDependent(precedentArea.Area, dependent);
            }
        }

        return precedents;
    }

    /// <summary>
    /// Remove formula from the dependency tree.
    /// </summary>
    /// <param name="formula">Formula to remove.</param>
    internal void RemoveFormula(XLCellFormula formula)
    {
        if (!this._dependencies.TryGetValue(formula, out FormulaDependencies? dependencies))
        {
            return;
        }

        this._dependencies.Remove(formula);
        foreach (SheetArea precedentArea in dependencies.Areas)
        {
            if (
                !this._sheetTrees.TryGetValue(
                    precedentArea.Name,
                    out SheetDependencyTree? sheetTree
                )
            )
            {
                throw new InvalidOperationException(
                    $"Dependency tree for sheet '{precedentArea.Name}' not found."
                );
            }

            sheetTree.RemoveDependent(precedentArea.Area, formula);
        }
    }

    internal void AddSheetTree(IXLWorksheet sheet) =>
        this._sheetTrees.Add(sheet.Name, new SheetDependencyTree());

    internal void RenameSheet(string oldSheetName, string newSheetName)
    {
        foreach (FormulaDependencies formulaDependencies in this._dependencies.Values)
        {
            formulaDependencies.RenameSheet(oldSheetName, newSheetName);
        }

        SheetDependencyTree renamedSheetTree = this._sheetTrees[oldSheetName];
        this._sheetTrees.Remove(oldSheetName);
        this._sheetTrees.Add(newSheetName, renamedSheetTree);

        foreach (SheetDependencyTree sheetTree in this._sheetTrees.Values)
        {
            sheetTree.RenameSheet(oldSheetName, newSheetName);
        }
    }

    /// <summary>
    /// Mark all formulas that depend (directly or transitively) on the area as dirty.
    /// </summary>
    internal void MarkDirty(SheetArea dirtyArea)
    {
        // BFS vs DFS: Although the longest chain found in the wild is 1000
        // formulas long, attacker could supply malicious excel with recursion
        // leading to stack overflow => use queue even with extra allocation cost.
        Queue<SheetArea> queue = new();
        queue.Enqueue(dirtyArea);
        while (queue.Count > 0)
        {
            SheetArea affectedArea = queue.Dequeue();
            SheetDependencyTree sheetTree = this._sheetTrees[affectedArea.Name];
            foreach (AreaDependents area in sheetTree.FindDependentsAreas(affectedArea.Area))
            {
                foreach (Dependent dependent in area.Dependents)
                {
                    // Ensure we don't end up in an infinite cycle
                    if (dependent.IsDirty)
                    {
                        continue;
                    }

                    dependent.MarkDirty();
                    queue.Enqueue(dependent.FormulaArea);
                }
            }
        }
    }

    private FormulaDependencies GetFormulaPrecedents(
        SheetArea formulaArea,
        XLCellFormula formula,
        XLWorkbook workbook
    )
    {
        Formula ast = formula.GetAst(workbook.CalcEngine);
        DependenciesContext context = new(formulaArea, workbook);
        List<SheetArea>? rootReference = ast.AstRoot.Accept(context, this._visitor);

        // If formula references are propagated to the root, make sure to add them.
        if (rootReference is not null)
        {
            context.AddAreas(rootReference);
        }

        return context.Dependencies;
    }

    /// <summary>
    /// An area that is referred by formulas in different cells, i.e. it
    /// contains precedent cells for a formula. If anything in the area
    /// potentially changes, all dependents might also change.
    /// </summary>
    private class AreaDependents : ISpatialData
    {
        /// <summary>
        /// An area in a sheet that is used by formulas, converted to RBush envelope.
        /// All RBush <c>double</c> coordinates are whole numbers.
        /// </summary>
        private readonly Envelope _area;

        private readonly List<Dependent> _dependents;

        internal AreaDependents(in Envelope area, Dependent firstDependent)
        {
            this._area = area;
            this._dependents = [firstDependent];
        }

        /// <summary>
        /// The area in a sheet on which some formulas depend on.
        /// </summary>
        /// <example><c>SIN(A4)</c> depends on <c>A4:A4</c> area.</example>.
        public ref readonly Envelope Envelope => ref this._area;

        /// <summary>
        /// List of formulas that depend on the range, always at least one.
        /// </summary>
        internal IReadOnlyList<Dependent> Dependents => this._dependents;

        internal void AddDependent(Dependent dependent) => this._dependents.Add(dependent);

        internal void RemoveDependent(XLCellFormula formula)
        {
            for (int i = 0; i < this._dependents.Count; ++i)
            {
                Dependent dependent = this._dependents[i];

                // several different formulas can depend on same area,
                // remove only dependent of the formula.
                if (dependent.Formula != formula)
                {
                    continue;
                }

                // Remove from list by moving the last element to the removed
                // element place and decrease capacity.
                this._dependents[i] = this._dependents[this._dependents.Count - 1];

                // Remove last item, capacity is unchanged, only list size is updated.
                this._dependents.RemoveAt(this._dependents.Count - 1);
            }
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            for (int i = 0; i < this._dependents.Count; ++i)
            {
                Dependent dependent = this._dependents[i];
                if (
                    XlsxSharp.XLHelper.SheetComparer.Equals(
                        dependent.FormulaArea.Name,
                        oldSheetName
                    )
                )
                {
                    SheetArea renamedArea = new(newSheetName, dependent.FormulaArea.Area);
                    this._dependents[i] = new Dependent(renamedArea, dependent.Formula);
                }
            }
        }
    }

    /// <summary>
    /// A dependent on a precedent area. If the precedent area changes,
    /// the dependent might also now be invalid.
    /// </summary>
    private readonly struct Dependent
    {
        /// <summary>
        /// Area that is invalidated, when precedent area is marked as
        /// dirty. Generally, it is an area of formula (1x1 for normal
        /// formulas), larger for array formulas. Cell formula by itself
        /// doesn't contain it's address to make it easier add/delete
        /// rows/cols.
        /// </summary>
        internal readonly SheetArea FormulaArea;

        internal Dependent(SheetArea formulaArea, XLCellFormula formula)
        {
            this.FormulaArea = formulaArea;
            this.Formula = formula;
        }

        /// <summary>
        /// The formula that is affected by changes in precedent area.
        /// </summary>
        internal XLCellFormula Formula { get; }

        internal bool IsDirty => this.Formula.IsDirty;

        internal bool MarkDirty() => this.Formula.IsDirty = true;
    }

    /// <summary>
    /// A dependency tree for a single worksheet.
    /// </summary>
    private class SheetDependencyTree
    {
        /// <summary>
        /// The precedent areas are not duplicated, though two areas might overlap.
        /// </summary>
        private readonly RBush<AreaDependents> _tree;

        /// <summary>
        /// All precedent areas in the sheet for all formulas in the workbook.
        /// </summary>
        /// <remarks>
        /// Not sure extra memory (at least 32 bytes per formula) is worth less CPU: O(1) vs O(log N)....
        /// </remarks>
        private readonly Dictionary<Area, AreaDependents> _precedentAreas;

        internal SheetDependencyTree()
        {
            this._tree = new RBush<AreaDependents>();
            this._precedentAreas = new Dictionary<Area, AreaDependents>();
        }

        internal bool IsEmpty => this._tree.Count == 0;

        internal void AddDependent(Area precedentRange, Dependent dependent)
        {
            if (
                !this._precedentAreas.TryGetValue(precedentRange, out AreaDependents? precedentArea)
            )
            {
                precedentArea = new AreaDependents(ToEnvelope(precedentRange), dependent);
                this._precedentAreas.Add(precedentRange, precedentArea);
                this._tree.Insert(precedentArea);
            }
            else
            {
                precedentArea.AddDependent(dependent);
            }
        }

        internal IReadOnlyList<AreaDependents> FindDependentsAreas(Area dirtyRange) =>
            this._tree.Search(ToEnvelope(dirtyRange));

        /// <summary>
        /// Remove a dependency of <paramref name="formula"/> on a
        /// <paramref name="precedentRange"/> from the sheet dependency tree.
        /// </summary>
        /// <param name="precedentRange">A precedent area in the sheet.</param>
        /// <param name="formula">Formula depending on the <paramref name="precedentRange"/>.</param>
        internal void RemoveDependent(Area precedentRange, XLCellFormula formula)
        {
            if (
                !this._precedentAreas.TryGetValue(precedentRange, out AreaDependents? precedentArea)
            )
            {
                return;
            }

            precedentArea.RemoveDependent(formula);
            if (precedentArea.Dependents.Count == 0)
            {
                this._tree.Delete(precedentArea);
                this._precedentAreas.Remove(precedentRange);
            }
        }

        internal void RenameSheet(string oldSheetName, string newSheetName)
        {
            // Area dependents instances are shared among _precedentAreas and _tree, so it is
            // enough to change _precedentAreas.
            foreach (AreaDependents areaDependents in this._precedentAreas.Values)
            {
                areaDependents.RenameSheet(oldSheetName, newSheetName);
            }
        }

        private static Envelope ToEnvelope(Area range) =>
            new(range.LeftColumn, range.TopRow, range.RightColumn, range.BottomRow);
    }
}

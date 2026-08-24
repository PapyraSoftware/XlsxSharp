namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// A list of objects a cell formula depends on. If one of them changes,
/// the formula value might no longer be accurate and needs to be recalculated.
/// </summary>
internal class FormulaDependencies
{
    private readonly HashSet<SheetArea> _areas = [];
    private readonly HashSet<XLName> _names = [];

    /// <summary>
    /// List of areas the formula depends on. It is likely a superset of accurate
    /// result for unusual formulas, but if a value in an areas changes, the dependent
    /// formula should be marked as dirty.
    /// </summary>
    public IReadOnlyCollection<SheetArea> Areas => this._areas;

    /// <summary>
    /// A collection of names in the formula. If a name changes (added, deleted),
    /// the formula dependencies should be refreshed, because new name might refer to
    /// different references (e.g. a name previously referred to <c>A5</c> and is redefined
    /// to <c>B7</c> or just value <c>7</c> =&gt; formula no longer depends on <c>A5</c>).
    /// </summary>
    public IReadOnlyCollection<XLName> Names => this._names;

    internal void AddAreas(List<SheetArea> sheetAreas) => this._areas.UnionWith(sheetAreas);

    internal void AddName(XLName name) => this._names.Add(name);

    internal void RenameSheet(string oldSheetName, string newSheetName)
    {
        // The renaming is done for every formula, so only allocate when needed.
        List<(SheetArea Original, SheetArea Replacement)>? areasToRename = null;
        foreach (SheetArea areaInFormula in this._areas)
        {
            if (XlsxSharp.XLHelper.SheetComparer.Equals(areaInFormula.Name, oldSheetName))
            {
                SheetArea renamedArea = new(newSheetName, areaInFormula.Area);
                areasToRename ??= [];
                areasToRename.Add((areaInFormula, renamedArea));
            }
        }

        if (areasToRename is not null)
        {
            foreach ((SheetArea original, SheetArea replacement) in areasToRename)
            {
                this._areas.Remove(original);
                this._areas.Add(replacement);
            }
        }

        List<(XLName Original, XLName Replacement)>? namesToRename = null;
        foreach (XLName nameInFormula in this._names)
        {
            if (
                nameInFormula.SheetName is not null
                && XlsxSharp.XLHelper.SheetComparer.Equals(nameInFormula.SheetName, oldSheetName)
            )
            {
                XLName renamedName = new(newSheetName, nameInFormula.Name);
                namesToRename ??= [];
                namesToRename.Add((nameInFormula, renamedName));
            }
        }

        if (namesToRename is not null)
        {
            foreach ((XLName original, XLName replacement) in namesToRename)
            {
                this._names.Remove(original);
                this._names.Add(replacement);
            }
        }
    }
}

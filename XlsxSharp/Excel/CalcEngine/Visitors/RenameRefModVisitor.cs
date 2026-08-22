using System.Collections.Generic;
using System.Linq;
using ClosedXML.Parser;

namespace XlsxSharp.Excel.CalcEngine.Visitors;

/// <summary>
/// A factory to rename named reference object (sheets, tables ect.).
/// </summary>
internal class RenameRefModVisitor : RefModVisitor
{
    private readonly Dictionary<string, string?>? _sheets;
    private readonly Dictionary<string, string>? _tables;

    /// <summary>
    /// A mapping of sheets, from old name (key) to a new name (value).
    /// The <c>null</c> value indicates sheet has been deleted.
    /// </summary>
    internal IReadOnlyDictionary<string, string?> Sheets
    {
        init =>
            this._sheets = value.ToDictionary(
                x => x.Key,
                x => x.Value,
                XlsxSharp.XLHelper.SheetComparer
            );
    }

    internal IReadOnlyDictionary<string, string> Tables
    {
        init =>
            this._tables = value.ToDictionary(
                x => x.Key,
                x => x.Value,
                XlsxSharp.XLHelper.NameComparer
            );
    }

    protected override string? ModifySheet(ModContext ctx, string sheetName)
    {
        if (this._sheets is not null && this._sheets.TryGetValue(sheetName, out string? newName))
        {
            return newName;
        }

        return sheetName;
    }

    protected override string? ModifyTable(ModContext ctx, string tableName)
    {
        if (this._tables is not null && this._tables.TryGetValue(tableName, out string? newName))
        {
            return newName;
        }

        return tableName;
    }
}

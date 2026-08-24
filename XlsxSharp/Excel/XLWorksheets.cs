using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using XlsxSharp.Excel.Tables;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLWorksheets : IXLWorksheets, IEnumerable<XLWorksheet>
{
    private readonly XLWorkbook _workbook;
    private readonly Dictionary<string, XLWorksheet> _worksheets = new(
        XlsxSharp.XLHelper.SheetComparer
    );
    internal ICollection<string> Deleted { get; private set; }

    /// <summary>
    /// SheetId that will be assigned to next created sheet.
    /// </summary>
    private uint _nextSheetId = 1;

    #region Constructor

    public XLWorksheets(XLWorkbook workbook)
    {
        this._workbook = workbook;
        this.Deleted = (HashSet<string>)[];
    }

    #endregion Constructor

    #region IEnumerable<XLWorksheet> Members

    public IEnumerator<XLWorksheet> GetEnumerator() =>
        ((IEnumerable<XLWorksheet>)this._worksheets.Values).GetEnumerator();

    #endregion IEnumerable<XLWorksheet> Members

    #region IXLWorksheets Members

    public int Count
    {
        [DebuggerStepThrough]
        get => this._worksheets.Count;
    }

    public bool Contains(string sheetName) => this._worksheets.ContainsKey(sheetName);

    bool IXLWorksheets.TryGetWorksheet(
        string sheetName,
        [NotNullWhen(true)] out IXLWorksheet? worksheet
    )
    {
        if (this.TryGetWorksheet(sheetName, out XLWorksheet? foundSheet))
        {
            worksheet = foundSheet;
            return true;
        }

        worksheet = null;
        return false;
    }

    internal bool TryGetWorksheet(
        string sheetName,
        [NotNullWhen(true)] out XLWorksheet? worksheet
    ) => this._worksheets.TryGetValue(sheetName, out worksheet);

    public IXLWorksheet Worksheet(string sheetName)
    {
        if (!this._worksheets.TryGetValue(sheetName, out XLWorksheet? foundSheet))
        {
            throw new KeyNotFoundException($"There isn't a worksheet named '{sheetName}'.");
        }

        return foundSheet;
    }

    public IXLWorksheet Worksheet(int position)
    {
        int wsCount = this._worksheets.Values.Count(w => w.Position == position);
        if (wsCount == 0)
        {
            throw new ArgumentException("There isn't a worksheet associated with that position.");
        }

        if (wsCount > 1)
        {
            throw new ArgumentException(
                "Can't retrieve a worksheet because there are multiple worksheets associated with that position."
            );
        }

        return this._worksheets.Values.Single(w => w.Position == position);
    }

    public IXLWorksheet Add() => this.Add(this.GetNextWorksheetName());

    public IXLWorksheet Add(int position) => this.Add(this.GetNextWorksheetName(), position);

    public IXLWorksheet Add(string sheetName)
    {
        XLWorksheet sheet = new(sheetName, this._workbook, this.GetNextSheetId());
        this.Add(sheetName, sheet);
        sheet._position = this._worksheets.Count + this._workbook.UnsupportedSheets.Count;
        return sheet;
    }

    public IXLWorksheet Add(string sheetName, int position) =>
        this.Add(sheetName, position, this.GetNextSheetId());

    internal XLWorksheet Add(string sheetName, int position, uint sheetId)
    {
        this._worksheets.Values.Where(w => w._position >= position).ForEach(w => w._position += 1);
        this._workbook.UnsupportedSheets.Where(w => w.Position >= position)
            .ForEach(w => w.Position += 1);

        // If the loaded sheetId is greater than current, just make sure our next sheetId is even bigger.
        this._nextSheetId = Math.Max(this._nextSheetId, sheetId + 1);
        XLWorksheet sheet = new(sheetName, this._workbook, sheetId);
        this.Add(sheetName, sheet);
        sheet._position = position;
        return sheet;
    }

    private void Add(string sheetName, XLWorksheet sheet)
    {
        if (this._worksheets.ContainsKey(sheetName))
        {
            throw new ArgumentException(
                string.Format(
                    "A worksheet with the same name ({0}) has already been added.",
                    sheetName
                ),
                nameof(sheetName)
            );
        }

        this._worksheets.Add(sheetName, sheet);

        this._workbook.NotifyWorksheetAdded(sheet);
    }

    public void Delete(string sheetName) => this.Delete(this._worksheets[sheetName].Position);

    public void Delete(int position)
    {
        int wsCount = this._worksheets.Values.Count(w => w.Position == position);
        if (wsCount == 0)
        {
            throw new ArgumentException("There isn't a worksheet associated with that index.");
        }

        if (wsCount > 1)
        {
            throw new ArgumentException(
                "Can't delete the worksheet because there are multiple worksheets associated with that index."
            );
        }

        XLWorksheet ws = this._worksheets.Values.Single(w => w.Position == position);
        if (!string.IsNullOrWhiteSpace(ws.RelId) && !this.Deleted.Contains(ws.RelId))
        {
            this.Deleted.Add(ws.RelId);
        }

        this._worksheets.RemoveAll(w => w.Position == position);
        this._worksheets.Values.Where(w => w.Position > position).ForEach(w => w._position -= 1);
        this._workbook.UnsupportedSheets.Where(w => w.Position > position)
            .ForEach(w => w.Position -= 1);

        ws.Cleanup();
    }

    IEnumerator<IXLWorksheet> IEnumerable<IXLWorksheet>.GetEnumerator() =>
        this._worksheets.Values.Cast<IXLWorksheet>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IXLWorksheet Add(DataTable dataTable) => this.Add(dataTable, dataTable.TableName);

    public IXLWorksheet Add(DataTable dataTable, string sheetName) =>
        this.Add(dataTable, sheetName, TableNameGenerator.GetNewTableName(this._workbook));

    public IXLWorksheet Add(DataTable dataTable, string sheetName, string tableName)
    {
        IXLWorksheet ws = this.Add(sheetName);
        ws.Cell(1, 1).InsertTable(dataTable, tableName);
        return ws;
    }

    public void Add(DataSet dataSet)
    {
        foreach (DataTable t in dataSet.Tables)
        {
            this.Add(t);
        }
    }

    #endregion IXLWorksheets Members

    public void Rename(string oldSheetName, string newSheetName)
    {
        if (
            string.IsNullOrWhiteSpace(oldSheetName)
            || !this._worksheets.TryGetValue(oldSheetName, out XLWorksheet ws)
        )
        {
            return;
        }

        if (
            !oldSheetName.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)
            && this._worksheets.ContainsKey(newSheetName)
        )
        {
            throw new ArgumentException(
                string.Format(
                    "A worksheet with the same name ({0}) has already been added.",
                    newSheetName
                ),
                nameof(newSheetName)
            );
        }

        this._worksheets.Remove(oldSheetName);
        this.Add(newSheetName, ws);

        foreach (IWorkbookListener listener in this.GetWorkbookListeners())
        {
            listener.OnSheetRenamed(oldSheetName, newSheetName);
        }
    }

    #region Private members

    private IEnumerable<IWorkbookListener> GetWorkbookListeners()
    {
        // All components that should be updated when sheet is added/removed or renamed should
        // be enumerated here.
        yield return this._workbook.CalcEngine;

        foreach (XLWorksheet sheet in this._worksheets.Values)
        {
            yield return sheet.Internals.CellsCollection;
        }

        foreach (XLDefinedName definedName in this._workbook.DefinedNamesInternal)
        {
            yield return definedName;
        }

        foreach (XLWorksheet sheet in this._worksheets.Values)
        {
            foreach (XLDefinedName definedName in sheet.DefinedNames)
            {
                yield return definedName;
            }
        }
    }

    private string GetNextWorksheetName()
    {
        int worksheetNumber = this.Count + 1;
        string sheetName = $"Sheet{worksheetNumber}";
        while (
            this._worksheets.Values.Any(p =>
                p.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)
            )
        )
        {
            worksheetNumber++;
            sheetName = $"Sheet{worksheetNumber}";
        }
        return sheetName;
    }

    private uint GetNextSheetId() => this._nextSheetId++;

    #endregion Private members
}

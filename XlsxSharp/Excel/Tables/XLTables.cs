#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Tables;

internal class XLTables : IXLTables, IEnumerable<XLTable>
{
    private readonly Dictionary<String, XLTable> _tables;

    public XLTables()
    {
        this._tables = new Dictionary<String, XLTable>(StringComparer.OrdinalIgnoreCase);
        this.Deleted = (HashSet<string>)[];
    }

    internal ICollection<String> Deleted { get; }

    #region IXLTables Members

    bool IXLTables.TryGetTable(string tableName, out IXLTable table)
    {
        if (this.TryGetTable(tableName, out XLTable foundTable))
        {
            table = foundTable;
            return true;
        }

        table = default;
        return false;
    }

    public void Add(IXLTable table)
    {
        XLTable xlTable = (XLTable)table;
        this._tables.Add(table.Name, xlTable);
        xlTable.OnAddedToTables();
    }

    public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        this._tables.Values.ForEach(t => t.Clear(clearOptions));
        return this;
    }

    public Boolean Contains(String name) => this._tables.ContainsKey(name);

    public Dictionary<string, XLTable>.ValueCollection.Enumerator GetEnumerator() =>
        this._tables.Values.GetEnumerator();

    IEnumerator<XLTable> IEnumerable<XLTable>.GetEnumerator() => this.GetEnumerator();

    IEnumerator<IXLTable> IEnumerable<IXLTable>.GetEnumerator() => this.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void Remove(Int32 index) => this.Remove(this._tables.ElementAt(index).Key);

    public void Remove(String name)
    {
        if (!this._tables.TryGetValue(name, out XLTable table))
        {
            throw new ArgumentOutOfRangeException(
                nameof(name),
                $"Unable to delete table because the table name {name} could not be found."
            );
        }

        this._tables.Remove(name);

        string relId = table.RelId;

        if (relId is not null)
        {
            this.Deleted.Add(relId);
        }
    }

    public IXLTable Table(Int32 index) => this._tables.ElementAt(index).Value;

    public IXLTable Table(String name)
    {
        if (this.TryGetTable(name, out XLTable table))
        {
            return table;
        }

        throw new ArgumentOutOfRangeException(nameof(name), $"Table {name} was not found.");
    }

    internal bool TryGetTable(string tableName, out XLTable table) =>
        this._tables.TryGetValue(tableName, out table);

    #endregion IXLTables Members
}

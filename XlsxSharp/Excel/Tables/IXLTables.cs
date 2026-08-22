#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.Tables;

public interface IXLTables : IEnumerable<IXLTable>
{
    public void Add(IXLTable table);

    /// <summary>
    /// Clears the contents of these tables.
    /// </summary>
    /// <param name="clearOptions">Specify what you want to clear.</param>
    public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All);

    public bool Contains(string name);

    public void Remove(int index);

    public void Remove(string name);

    public IXLTable Table(int index);

    public IXLTable Table(string name);

    public bool TryGetTable(string tableName, out IXLTable table);
}

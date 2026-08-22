#nullable disable

using System.Collections.Generic;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Misc;

internal class XLIdManager
{
    private HashSet<int> _hash = [];

    public int GetNext()
    {
        if (this._hash.Count == 0)
        {
            this._hash.Add(1);
            return 1;
        }

        int id = 1;
        while (true)
        {
            if (!this._hash.Contains(id))
            {
                this._hash.Add(id);
                return id;
            }
            id++;
        }
    }

    public void Add(int value) => this._hash.Add(value);

    public void Add(IEnumerable<int> values) => values.ForEach(v => this._hash.Add(v));
}

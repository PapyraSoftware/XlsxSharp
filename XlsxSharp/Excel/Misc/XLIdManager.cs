#nullable disable

using System;
using System.Collections.Generic;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Misc;

internal class XLIdManager
{
    private HashSet<Int32> _hash = [];

    public Int32 GetNext()
    {
        if (this._hash.Count == 0)
        {
            this._hash.Add(1);
            return 1;
        }

        Int32 id = 1;
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

    public void Add(Int32 value) => this._hash.Add(value);

    public void Add(IEnumerable<Int32> values) => values.ForEach(v => this._hash.Add(v));
}

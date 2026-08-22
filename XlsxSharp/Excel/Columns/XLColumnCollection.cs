#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal class XLColumnsCollection : IDictionary<Int32, XLColumn>
{
    private readonly Dictionary<Int32, XLColumn> _dictionary = new();

    public void ShiftColumnsRight(Int32 startingColumn, Int32 columnsToShift)
    {
        foreach (
            int co in this
                ._dictionary.Keys.Where(k => k >= startingColumn)
                .OrderByDescending(k => k)
        )
        {
            XLColumn columnToMove = this._dictionary[co];
            this._dictionary.Remove(co);
            Int32 newColumnNum = co + columnsToShift;
            if (newColumnNum <= XlsxSharp.XLHelper.MaxColumnNumber)
            {
                columnToMove.SetColumnNumber(newColumnNum);
                this._dictionary.Add(newColumnNum, columnToMove);
            }
        }
    }

    public void Add(int key, XLColumn value)
    {
        this._dictionary.Add(key, value);
    }

    public bool ContainsKey(int key) => this._dictionary.ContainsKey(key);

    public ICollection<int> Keys => this._dictionary.Keys;

    public bool Remove(int key)
    {
        return this._dictionary.Remove(key);
    }

    public bool TryGetValue(int key, out XLColumn value)
    {
        return this._dictionary.TryGetValue(key, out value);
    }

    public ICollection<XLColumn> Values => this._dictionary.Values;

    public XLColumn this[int key]
    {
        get => this._dictionary[key];
        set => this._dictionary[key] = value;
    }

    public void Add(KeyValuePair<int, XLColumn> item)
    {
        this._dictionary.Add(item.Key, item.Value);
    }

    public void Clear()
    {
        this._dictionary.Clear();
    }

    public bool Contains(KeyValuePair<int, XLColumn> item)
    {
        return this._dictionary.Contains(item);
    }

    public void CopyTo(KeyValuePair<int, XLColumn>[] array, int arrayIndex)
    {
        throw new NotImplementedException();
    }

    public int Count => this._dictionary.Count;

    public bool IsReadOnly => false;

    public bool Remove(KeyValuePair<int, XLColumn> item)
    {
        return this._dictionary.Remove(item.Key);
    }

    public IEnumerator<KeyValuePair<int, XLColumn>> GetEnumerator() =>
        this._dictionary.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this._dictionary.GetEnumerator();

    public void RemoveAll(Func<XLColumn, Boolean> predicate)
    {
        this._dictionary.RemoveAll(predicate);
    }
}

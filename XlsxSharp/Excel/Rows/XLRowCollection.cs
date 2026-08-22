#nullable disable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Rows;

internal class XLRowsCollection : IDictionary<int, XLRow>
{
    private readonly Dictionary<int, XLRow> _dictionary = new();

    public Dictionary<int, XLRow> Deleted { get; } = new();

    public int MaxRowUsed;

    #region IDictionary<int,XLRow> Members

    public void Add(int key, XLRow value)
    {
        if (key > this.MaxRowUsed)
        {
            this.MaxRowUsed = key;
        }

        this.Deleted.Remove(key);
        this._dictionary.Add(key, value);
    }

    public bool ContainsKey(int key) => this._dictionary.ContainsKey(key);

    public ICollection<int> Keys => this._dictionary.Keys;

    public bool Remove(int key)
    {
        if (!this.Deleted.ContainsKey(key))
        {
            this.Deleted.Add(key, this._dictionary[key]);
        }

        return this._dictionary.Remove(key);
    }

    public bool TryGetValue(int key, out XLRow value) =>
        this._dictionary.TryGetValue(key, out value);

    public ICollection<XLRow> Values => this._dictionary.Values;

    public XLRow this[int key]
    {
        get => this._dictionary[key];
        set => this._dictionary[key] = value;
    }

    public void Add(KeyValuePair<int, XLRow> item)
    {
        if (item.Key > this.MaxRowUsed)
        {
            this.MaxRowUsed = item.Key;
        }

        this.Deleted.Remove(item.Key);
        this._dictionary.Add(item.Key, item.Value);
    }

    public void Clear() => this._dictionary.Clear();

    public bool Contains(KeyValuePair<int, XLRow> item) => this._dictionary.Contains(item);

    public void CopyTo(KeyValuePair<int, XLRow>[] array, int arrayIndex) =>
        throw new NotImplementedException();

    public int Count => this._dictionary.Count;

    public bool IsReadOnly => false;

    public bool Remove(KeyValuePair<int, XLRow> item)
    {
        if (!this.Deleted.ContainsKey(item.Key))
        {
            this.Deleted.Add(item.Key, this._dictionary[item.Key]);
        }

        return this._dictionary.Remove(item.Key);
    }

    public IEnumerator<KeyValuePair<int, XLRow>> GetEnumerator() =>
        this._dictionary.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this._dictionary.GetEnumerator();

    #endregion IDictionary<int,XLRow> Members

    public void ShiftRowsDown(int startingRow, int rowsToShift)
    {
        foreach (
            int ro in this._dictionary.Keys.Where(k => k >= startingRow).OrderByDescending(k => k)
        )
        {
            XLRow rowToMove = this._dictionary[ro];
            this._dictionary.Remove(ro);
            int newRowNum = ro + rowsToShift;
            if (newRowNum <= XlsxSharp.XLHelper.MaxRowNumber)
            {
                rowToMove.SetRowNumber(newRowNum);
                this._dictionary.Add(newRowNum, rowToMove);
            }
        }
    }

    public void RemoveAll(Func<XLRow, bool> predicate)
    {
        foreach (
            XLRow row in this
                ._dictionary.Values.Where(predicate)
                .Where(row1 => !this.Deleted.ContainsKey(row1.RowNumber()))
        )
        {
            this.Deleted.Add(row.RowNumber(), row);
        }

        this._dictionary.RemoveAll(predicate);
    }
}

#nullable disable

using System;
using System.Collections.Generic;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel.Misc;

public class XLDictionary<T> : Dictionary<int, T>
{
    public XLDictionary() { }

    public XLDictionary(XLDictionary<T> other) => other.Values.ForEach(this.Add);

    public void Initialize(T value)
    {
        if (this.Count > 0)
        {
            this.Clear();
        }

        this.Add(value);
    }

    public void Add(T value) => this.Add(this.Count + 1, value);

    internal XLDictionary<T> CopyDictionary()
    {
        XLDictionary<T> copy = new();
        foreach ((int key, T value) in this)
        {
            copy.Add(key, value);
        }

        return copy;
    }
}

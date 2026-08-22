#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Extensions;

internal static class DictionaryExtensions
{
    public static void RemoveAll<TKey, TValue>(
        this Dictionary<TKey, TValue> dic,
        Func<TValue, bool> predicate
    )
    {
        List<TKey> keys = [.. dic.Keys.Where(k => predicate(dic[k]))];
        foreach (TKey key in keys)
        {
            dic.Remove(key);
        }
    }
}

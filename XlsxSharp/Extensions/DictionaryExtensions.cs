#nullable disable

namespace XlsxSharp.Extensions;

internal static class DictionaryExtensions
{
    public static void RemoveAll<TKey, TValue>(
        this Dictionary<TKey, TValue> dic,
        Func<TValue, bool> predicate
    )
    {
        List<TKey> keys = [.. dic.Where(kvp => predicate(kvp.Value)).Select(kvp => kvp.Key)];
        foreach (TKey key in keys)
        {
            dic.Remove(key);
        }
    }
}

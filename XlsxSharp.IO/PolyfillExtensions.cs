namespace XlsxSharp.IO;

internal static class PolyfillExtensions
{
#if !NET5_0_OR_GREATER
    public static bool TryAdd<TKey, TValue>(
        this Dictionary<TKey, TValue> dictionary,
        TKey key,
        TValue value
    )
        where TKey : notnull
    {
        if (!dictionary.ContainsKey(key))
        {
            dictionary.Add(key, value);
            return true;
        }

        return false;
    }
#endif
}

#if !NET5_0_OR_GREATER
internal sealed class ReferenceEqualityComparer : IEqualityComparer<object?>, IEqualityComparer
{
    public static readonly ReferenceEqualityComparer Instance = new();

    private ReferenceEqualityComparer() { }

    public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);

    public int GetHashCode(object? obj) => RuntimeHelpers.GetHashCode(obj);
}
#endif

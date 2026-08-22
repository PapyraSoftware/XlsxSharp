using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Utils;

/// <summary>
/// A dictionary that can find value for a key, but also key for a value. The important property is
/// that while the key must be unique, but value might be duplicate.
/// <example>
/// An example is a font list loaded from the file. The id is a <c>ST_FontId</c> and values in
/// the list can be duplicate (two same fonts with different id).
/// </example>
/// </summary>
internal class BiDictionary<TKey, TValue> : IReadOnlyBiDictionary<TKey, TValue>
    where TKey : notnull
    where TValue : notnull
{
    private readonly Dictionary<TKey, TValue> _keyToValue;

    /// <summary>
    /// A reverse dictionary. The original <see cref="_keyToValue"/> can contain same entry multiple times.
    /// </summary>
    private readonly Dictionary<TValue, TKey> _entryToKey;

    internal BiDictionary()
    {
        this._entryToKey = new Dictionary<TValue, TKey>();
        this._keyToValue = new Dictionary<TKey, TValue>();
    }

    internal BiDictionary(int capacity)
    {
        this._entryToKey = new Dictionary<TValue, TKey>(capacity);
        this._keyToValue = new Dictionary<TKey, TValue>(capacity);
    }

    public TValue this[TKey key] => this._keyToValue[key];

    public IEnumerable<TKey> Keys => this._keyToValue.Keys;

    public IEnumerable<TValue> Values => this._keyToValue.Values;

    public TKey this[TValue value] => this._entryToKey[value];

    public int Count => this._keyToValue.Count;

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator() =>
        this._keyToValue.GetEnumerator();

    public bool TryGetValue(TKey key, [MaybeNullWhen(false)] out TValue value) =>
        this._keyToValue.TryGetValue(key, out value);

    public bool ContainsKey(TKey key) => this._keyToValue.ContainsKey(key);

    public bool ContainsValue(TValue value) => this._entryToKey.ContainsKey(value);

    internal IReadOnlyDictionary<TKey, TValue> KeyToValue => this._keyToValue;

    internal IReadOnlyDictionary<TValue, TKey> ValueToKey => this._entryToKey;

    public void Add(TKey id, TValue value)
    {
        this._keyToValue.Add(id, value);

        // Keep first one. Entries should be added in ascending order (or at least order from
        // a file) and we want to reuse the earliest one to make things predictable.
        this._entryToKey.TryAdd(value, id);
    }

    public bool TryGetValue(TValue value, [NotNullWhen(true)] out TValue? foundValue)
    {
        if (this._entryToKey.TryGetValue(value, out TKey? key))
        {
            foundValue = this._keyToValue[key];
            return true;
        }

        foundValue = default;
        return false;
    }
}

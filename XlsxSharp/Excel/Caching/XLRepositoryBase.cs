using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.Caching;

internal abstract class XLRepositoryBase : IXLRepository
{
    public abstract void Clear();
}

internal abstract class XLRepositoryBase<Tkey, Tvalue>
    : XLRepositoryBase,
        IXLRepository<Tkey, Tvalue>
    where Tkey : struct, IEquatable<Tkey>
    where Tvalue : class
{
    const int CONCURRENCY_LEVEL = 4;
    const int INITIAL_CAPACITY = 1000;

    private readonly ConcurrentDictionary<Tkey, WeakReference> _storage;
    private readonly Func<Tkey, Tvalue> _createNew;

    protected XLRepositoryBase(Func<Tkey, Tvalue> createNew)
        : this(createNew, EqualityComparer<Tkey>.Default) { }

    protected XLRepositoryBase(Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer)
    {
        this._storage = new ConcurrentDictionary<Tkey, WeakReference>(
            CONCURRENCY_LEVEL,
            INITIAL_CAPACITY,
            comparer
        );
        this._createNew = createNew;
    }

    /// <summary>
    /// Check if the specified key is presented in the repository.
    /// </summary>
    /// <param name="key">Key to look for.</param>
    /// <param name="value">Value from the repository stored under specified key or null if key does
    /// not exist or the entry under this key has already bee GCed.</param>
    /// <returns>True if entry exists and alive, false otherwise.</returns>
    public bool ContainsKey(ref Tkey key, out Tvalue? value)
    {
        if (this._storage.TryGetValue(key, out WeakReference? cachedReference))
        {
            value = cachedReference.Target as Tvalue;
            return value != null;
        }
        value = null;
        return false;
    }

    /// <summary>
    /// Put the entity into the repository under the specified key if no other entity with
    /// the same key is presented.
    /// </summary>
    /// <param name="key">Key to identify the entity.</param>
    /// <param name="value">Entity to store.</param>
    /// <returns>Entity that is stored in the repository under the specified key
    /// (it can be either the <paramref name="value"/> or another entity that has been added to
    /// the repository before.)</returns>
    public Tvalue? Store(ref Tkey key, Tvalue value)
    {
        if (value is null)
        {
            return null;
        }

        do
        {
            if (
                this._storage.TryGetValue(key, out WeakReference? cachedReference)
                && cachedReference.Target is Tvalue storedValue
            )
            {
                return storedValue;
            }
        } while (!this._storage.TryAdd(key, new WeakReference(value)));

        return value;
    }

    public Tvalue GetOrCreate(ref Tkey key)
    {
        if (
            this._storage.TryGetValue(key, out WeakReference? cachedReference)
            && cachedReference.Target is Tvalue storedValue
        )
        {
            return storedValue;
        }

        this._storage.TryRemove(key, out WeakReference _);
        Tvalue value = this._createNew(key);
        return this.Store(ref key, value)!;
    }

    public Tvalue? Replace(ref Tkey oldKey, ref Tkey newKey)
    {
        if (
            this._storage.TryRemove(oldKey, out WeakReference? cachedReference)
            && cachedReference != null
        )
        {
            this._storage.TryAdd(newKey, cachedReference);
            return this.GetOrCreate(ref newKey);
        }

        return null;
    }

    public void Remove(ref Tkey key) => this._storage.TryRemove(key, out WeakReference _);

    public override void Clear() => this._storage.Clear();

    /// <summary>
    /// Enumerate items in repository removing "dead" entries.
    /// </summary>
    public IEnumerator<Tvalue> GetEnumerator() =>
        this
            ._storage.Select(pair =>
            {
                Tvalue? val = pair.Value.Target as Tvalue;
                if (val == null)
                {
                    this._storage.TryRemove(pair.Key, out WeakReference _);
                }
                return val;
            })
            .Where(val => val != null)
            .GetEnumerator()!;

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();
}

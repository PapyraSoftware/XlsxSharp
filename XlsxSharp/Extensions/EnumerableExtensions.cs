using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Extensions;

internal static class EnumerableExtensions
{
    public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
    {
        foreach (T item in source)
        {
            action(item);
        }
    }

    public static Type? GetItemType(this IEnumerable source)
    {
        return GetGenericArgument(source.GetType());

        Type? GetGenericArgument(Type collectionType)
        {
            Type? ienumerable = collectionType
                .GetInterfaces()
                .SingleOrDefault(i =>
                    i.GetGenericArguments().Length == 1 && i.Name == "IEnumerable`1"
                );

            return ienumerable?.GetGenericArguments()?.FirstOrDefault();
        }
    }

    /// <summary>
    /// Skip last element of a sequence.
    /// </summary>
    public static IEnumerable<T> SkipLast<T>(this IEnumerable<T> source)
    {
        using IEnumerator<T> enumerator = source.GetEnumerator();
        if (!enumerator.MoveNext())
        {
            yield break;
        }

        T prev = enumerator.Current;
        while (enumerator.MoveNext())
        {
            yield return prev;
            prev = enumerator.Current;
        }
    }

    public static bool HasDuplicates<T>(this IEnumerable<T> source)
    {
        HashSet<T> distinctItems = [];
        foreach (T item in source)
        {
            if (!distinctItems.Add(item))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Select all <typeparamref name="TItem"/> that are not null.
    /// </summary>
    public static IEnumerable<TItem> WhereNotNull<T, TItem>(
        this IEnumerable<T> source,
        Func<T, TItem?> property
    )
        where TItem : struct =>
        source.Select(property).Where(x => x.HasValue).Select(x => x!.Value);
}

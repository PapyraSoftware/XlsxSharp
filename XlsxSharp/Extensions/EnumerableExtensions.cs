using System.Collections;

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

            return ienumerable?.GetGenericArguments().FirstOrDefault();
        }
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

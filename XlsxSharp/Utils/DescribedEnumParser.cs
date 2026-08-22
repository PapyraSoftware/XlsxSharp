#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace XlsxSharp.Utils;

internal static class DescribedEnumParser<T>
{
    private static Lazy<IDictionary<string, T>> fromDescriptions = new(() =>
    {
        return ParseEnumDescriptions().ToDictionary(a => a.Item2, a => a.Item1);
    });

    private static Lazy<IDictionary<T, string>> toDescriptions = new(() =>
    {
        return ParseEnumDescriptions().ToDictionary(a => a.Item1, a => a.Item2);
    });

    public static T FromDescription(String value) => fromDescriptions.Value[value];

    public static Boolean IsValidDescription(String value) =>
        fromDescriptions.Value.ContainsKey(value);

    public static String ToDescription(T value) => toDescriptions.Value[value];

    private static IEnumerable<Tuple<T, String>> ParseEnumDescriptions()
    {
        Type type = typeof(T);
        return type.GetFields(BindingFlags.Public | BindingFlags.Static)
            .Select(f =>
            {
                IEnumerable<DescriptionAttribute> attributes = f.GetCustomAttributes(
                        typeof(DescriptionAttribute),
                        inherit: false
                    )
                    .OfType<DescriptionAttribute>();
                string description = attributes.FirstOrDefault()?.Description ?? f.Name;
                return new Tuple<T, String>((T)Enum.Parse(type, f.Name), description);
            });
    }
}

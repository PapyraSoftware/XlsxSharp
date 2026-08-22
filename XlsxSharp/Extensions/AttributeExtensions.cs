#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using System.Reflection;

namespace XlsxSharp.Extensions;

internal static class AttributeExtensions
{
    public static TAttribute[] GetAttributes<TAttribute>(this MemberInfo member)
        where TAttribute : Attribute
    {
        object[] attributes = member.GetCustomAttributes(typeof(TAttribute), true);

        return (TAttribute[])attributes;
    }

    public static bool HasAttribute<TAttribute>(this MemberInfo member)
        where TAttribute : Attribute => GetAttributes<TAttribute>(member).Any();
}

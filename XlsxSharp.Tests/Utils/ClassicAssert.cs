using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using TUnit.Core;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// Matches NUnit's <c>TestDelegate</c> so existing <c>TestDelegate action = () =&gt; ...;</c>
/// declarations keep compiling unchanged after the move away from NUnit.
/// </summary>
public delegate void TestDelegate();

/// <summary>
/// Thrown by <see cref="ClassicAssert"/> members on failure. TUnit reports any exception that
/// escapes a test as a failure, so this needs no framework-specific base type.
/// </summary>
public sealed class AssertionException(string message) : Exception(message);

/// <summary>
/// A synchronous, NUnit-classic-shaped assertion surface (<c>Assert.AreEqual</c>, <c>Assert.IsTrue</c>, ...)
/// so the bulk of the test suite could move to TUnit via a mechanical rename of the call sites
/// (<c>Assert.</c> -&gt; <c>ClassicAssert.</c>) instead of a hand rewrite of every assertion.
/// </summary>
public static class ClassicAssert
{
    public static void AreEqual(
        object? expected,
        object? actual,
        string? message = null,
        params object?[]? args
    )
    {
        if (!ValuesEqual(expected, actual))
        {
            Fail(FormatEquality("equal", expected, actual, message, args));
        }
    }

    public static void AreEqual(
        double expected,
        double actual,
        double delta,
        string? message = null,
        params object?[]? args
    )
    {
        bool equal =
            double.IsNaN(expected) || double.IsNaN(actual)
                ? double.Equals(expected, actual)
                : Math.Abs(expected - actual) <= delta;

        if (!equal)
        {
            Fail(FormatEquality("equal", expected, actual, message, args));
        }
    }

    public static void AreNotEqual(
        object? expected,
        object? actual,
        string? message = null,
        params object?[]? args
    )
    {
        if (ValuesEqual(expected, actual))
        {
            Fail(FormatEquality("not equal", expected, actual, message, args));
        }
    }

    public static void AreSame(object? expected, object? actual, string? message = null)
    {
        if (!ReferenceEquals(expected, actual))
        {
            Fail(
                message
                    ?? $"Expected same reference. Expected: {Describe(expected)}, Actual: {Describe(actual)}"
            );
        }
    }

    public static void AreNotSame(object? expected, object? actual, string? message = null)
    {
        if (ReferenceEquals(expected, actual))
        {
            Fail(message ?? "Expected different references but both were the same instance.");
        }
    }

    public static void IsTrue(bool condition, string? message = null) => True(condition, message);

    public static void True(bool condition, string? message = null)
    {
        if (!condition)
        {
            Fail(message ?? "Expected: True but was: False");
        }
    }

    public static void IsFalse(bool condition, string? message = null) => False(condition, message);

    public static void False(bool condition, string? message = null)
    {
        if (condition)
        {
            Fail(message ?? "Expected: False but was: True");
        }
    }

    public static void IsTrue(bool? condition, string? message = null) =>
        True(condition ?? false, message);

    public static void True(bool? condition, string? message = null) =>
        True(condition ?? false, message);

    public static void IsFalse(bool? condition, string? message = null) =>
        False(condition ?? false, message);

    public static void False(bool? condition, string? message = null) =>
        False(condition ?? false, message);

    public static void IsNull(object? value, string? message = null) => Null(value, message);

    public static void Null(object? value, string? message = null)
    {
        if (value is not null)
        {
            Fail(message ?? $"Expected: null but was: {Describe(value)}");
        }
    }

    public static void IsNotNull(object? value, string? message = null) => NotNull(value, message);

    public static void NotNull(object? value, string? message = null)
    {
        if (value is null)
        {
            Fail(message ?? "Expected: not null but was: null");
        }
    }

    public static void IsInstanceOf<T>(object? value, string? message = null)
    {
        if (value is not T)
        {
            Fail(
                message
                    ?? $"Expected an instance of {typeof(T)} but was {value?.GetType().ToString() ?? "null"}"
            );
        }
    }

    public static void Greater(IComparable actual, IComparable expected, string? message = null)
    {
        if (Compare(actual, expected) <= 0)
        {
            Fail(
                message ?? $"Expected: greater than {Describe(expected)} but was {Describe(actual)}"
            );
        }
    }

    public static void IsEmpty(object value, string? message = null)
    {
        bool empty = value switch
        {
            string s => s.Length == 0,
            IEnumerable e => IsEnumerableEmpty(e),
            _ => throw new ArgumentException(
                $"Cannot determine emptiness of {value.GetType()}",
                nameof(value)
            ),
        };
        if (!empty)
        {
            Fail(message ?? $"Expected empty but was {Describe(value)}");
        }
    }

    public static void Fail(string? message = null) =>
        throw new AssertionException(message ?? "Assertion failed.");

    public static void Ignore(string? message = null) => Skip.Test(message ?? "Ignored");

    public static TActual Throws<TActual>(TestDelegate code, string? message = null)
        where TActual : Exception
    {
        try
        {
            code();
        }
        catch (Exception ex)
        {
            if (ex is TActual typed)
            {
                return typed;
            }

            Fail(
                message
                    ?? $"Expected exception of type {typeof(TActual)} but got {ex.GetType()}: {ex.Message}"
            );
        }

        Fail(message ?? $"Expected exception of type {typeof(TActual)} but none was thrown.");
        throw new InvalidOperationException(); // unreachable, Fail always throws
    }

    public static Exception Throws(Type expectedType, TestDelegate code, string? message = null)
    {
        try
        {
            code();
        }
        catch (Exception ex)
        {
            if (expectedType.IsInstanceOfType(ex))
            {
                return ex;
            }

            Fail(
                message
                    ?? $"Expected exception of type {expectedType} but got {ex.GetType()}: {ex.Message}"
            );
        }

        Fail(message ?? $"Expected exception of type {expectedType} but none was thrown.");
        throw new InvalidOperationException(); // unreachable, Fail always throws
    }

    public static void DoesNotThrow(
        TestDelegate code,
        string? message = null,
        params object?[]? args
    )
    {
        try
        {
            code();
        }
        catch (Exception ex)
        {
            string formatted =
                message is not null && args is { Length: > 0 }
                    ? string.Format(CultureInfo.InvariantCulture, message, args)
                    : message ?? $"Expected no exception but got {ex.GetType()}: {ex.Message}";
            Fail(formatted);
        }
    }

    internal static bool ValuesEqual(object? expected, object? actual)
    {
        if (ReferenceEquals(expected, actual))
        {
            return true;
        }

        if (expected is null || actual is null)
        {
            return false;
        }

        // NUnit's classic comparer treats any two IEnumerable operands (arrays, lists, ...) as
        // ordered sequences and compares them element-by-element, regardless of whether they're
        // the same concrete collection type. Do the same here, ahead of the same-type shortcut
        // below: two different array instances with equal elements are otherwise only
        // reference-equal via Array's default Equals, which would wrongly fail.
        if (
            expected is not string
            && actual is not string
            && expected is IEnumerable expectedEnumerable
            && actual is IEnumerable actualEnumerable
        )
        {
            return SequencesEqual(expectedEnumerable, actualEnumerable);
        }

        if (expected.GetType() == actual.GetType())
        {
            return expected.Equals(actual);
        }

        if (IsNumeric(expected) && IsNumeric(actual))
        {
            return Convert
                .ToDouble(expected, CultureInfo.InvariantCulture)
                .Equals(Convert.ToDouble(actual, CultureInfo.InvariantCulture));
        }

        // NUnit's classic comparer falls back to a same-shaped IEquatable<T>/public bool
        // Equals(T) overload when the two operands have different runtime types (this is how,
        // e.g., XLCellValue.Equals(int) lets `Assert.AreEqual(1, cellValue)` work). Replicate
        // that via reflection instead of assuming both sides define a symmetric object.Equals.
        if (TryTypedEquals(actual, expected, out bool r1))
        {
            return r1;
        }

        if (TryTypedEquals(expected, actual, out bool r2))
        {
            return r2;
        }

        return expected.Equals(actual) || actual.Equals(expected);
    }

    private static bool SequencesEqual(IEnumerable expected, IEnumerable actual)
    {
        IEnumerator e1 = expected.GetEnumerator();
        IEnumerator e2 = actual.GetEnumerator();
        try
        {
            while (true)
            {
                bool has1 = e1.MoveNext();
                bool has2 = e2.MoveNext();
                if (has1 != has2)
                {
                    return false;
                }

                if (!has1)
                {
                    return true;
                }

                if (!ValuesEqual(e1.Current, e2.Current))
                {
                    return false;
                }
            }
        }
        finally
        {
            (e1 as IDisposable)?.Dispose();
            (e2 as IDisposable)?.Dispose();
        }
    }

    private static bool IsEnumerableEmpty(IEnumerable enumerable)
    {
        IEnumerator enumerator = enumerable.GetEnumerator();
        try
        {
            return !enumerator.MoveNext();
        }
        finally
        {
            (enumerator as IDisposable)?.Dispose();
        }
    }

    private static int Compare(object a, object b)
    {
        if (IsNumeric(a) && IsNumeric(b))
        {
            return Convert
                .ToDouble(a, CultureInfo.InvariantCulture)
                .CompareTo(Convert.ToDouble(b, CultureInfo.InvariantCulture));
        }

        return Comparer<object>.Default.Compare(a, b);
    }

    private static bool IsNumeric(object value) =>
        value
            is sbyte
                or byte
                or short
                or ushort
                or int
                or uint
                or long
                or ulong
                or float
                or double
                or decimal;

    private static readonly ConcurrentDictionary<
        (Type Instance, Type Argument),
        MethodInfo?
    > EquatableMethodCache = new();

    private static bool TryTypedEquals(object instance, object argument, out bool result)
    {
        (Type, Type) key = (instance.GetType(), argument.GetType());
        MethodInfo? method = EquatableMethodCache.GetOrAdd(key, FindEquatableMethod);

        if (method is null)
        {
            result = false;
            return false;
        }

        result = (bool)method.Invoke(instance, [argument])!;
        return true;
    }

    private static MethodInfo? FindEquatableMethod((Type Instance, Type Argument) k)
    {
        // A same-shaped public bool Equals(T) overload (e.g. XLCellValue.Equals(int)).
        MethodInfo? publicMethod = k
            .Instance.GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .FirstOrDefault(m =>
                m.Name == nameof(Equals)
                && m.ReturnType == typeof(bool)
                && !m.IsGenericMethod
                && m.GetParameters() is [{ ParameterType: var p }]
                && p == k.Argument
            );

        if (publicMethod is not null)
        {
            return publicMethod;
        }

        // An *explicit* IEquatable<T> implementation (e.g. StyleId's `bool
        // IEquatable<int>.Equals(int)`) isn't returned by GetMethods as a same-named public
        // method above; it's only reachable through the interface map.
        Type equatableInterface = typeof(IEquatable<>).MakeGenericType(k.Argument);
        if (!equatableInterface.IsAssignableFrom(k.Instance))
        {
            return null;
        }

        InterfaceMapping map = k.Instance.GetInterfaceMap(equatableInterface);
        int index = Array.IndexOf(
            map.InterfaceMethods,
            equatableInterface.GetMethod(nameof(Equals))
        );
        return index >= 0 ? map.TargetMethods[index] : null;
    }

    private static string Describe(object? value) =>
        value switch
        {
            null => "null",
            string s => $"\"{s}\"",
            _ => value.ToString() ?? value.GetType().ToString(),
        };

    private static string FormatEquality(
        string relation,
        object? expected,
        object? actual,
        string? message,
        object?[]? args
    )
    {
        string core = $"Expected {Describe(expected)} to be {relation} {Describe(actual)}.";
        if (message is null)
        {
            return core;
        }

        string formatted = args is { Length: > 0 }
            ? string.Format(CultureInfo.InvariantCulture, message, args)
            : message;
        return $"{core} {formatted}";
    }
}

/// <summary>
/// Mirrors NUnit's <c>CollectionAssert</c> surface used by this suite, on top of <see cref="ClassicAssert"/>.
/// </summary>
public static class CollectionAssert
{
    public static void AreEqual(IEnumerable expected, IEnumerable actual, string? message = null)
    {
        List<object?> expectedList = ToList(expected);
        List<object?> actualList = ToList(actual);
        bool equal =
            expectedList.Count == actualList.Count
            && expectedList.Zip(actualList, ClassicAssert.ValuesEqual).All(b => b);

        if (!equal)
        {
            ClassicAssert.Fail(
                message
                    ?? $"Expected collection [{Join(expectedList)}] but was [{Join(actualList)}]."
            );
        }
    }

    public static void AreEquivalent(
        IEnumerable expected,
        IEnumerable actual,
        string? message = null
    )
    {
        List<object?> remaining = ToList(actual);
        foreach (object? item in ToList(expected))
        {
            int index = remaining.FindIndex(x => ClassicAssert.ValuesEqual(item, x));
            if (index < 0)
            {
                ClassicAssert.Fail(
                    message
                        ?? $"Collections are not equivalent: missing {ClassicAssert_Describe(item)}."
                );
                return;
            }

            remaining.RemoveAt(index);
        }

        if (remaining.Count > 0)
        {
            ClassicAssert.Fail(
                message ?? $"Collections are not equivalent: extra [{Join(remaining)}]."
            );
        }
    }

    public static void IsEmpty(IEnumerable collection, string? message = null) =>
        ClassicAssert.IsEmpty(collection, message);

    private static List<object?> ToList(IEnumerable enumerable)
    {
        List<object?> list = [];
        foreach (object? item in enumerable)
        {
            list.Add(item);
        }

        return list;
    }

    private static string Join(IEnumerable<object?> items) =>
        string.Join(", ", items.Select(ClassicAssert_Describe));

    private static string ClassicAssert_Describe(object? value) =>
        value switch
        {
            null => "null",
            string s => $"\"{s}\"",
            _ => value.ToString() ?? value.GetType().ToString(),
        };
}

/// <summary>
/// Mirrors NUnit's <c>StringAssert</c> surface used by this suite.
/// </summary>
public static class StringAssert
{
    public static void Contains(string expectedSubstring, string? actual, string? message = null)
    {
        if (actual is null || !actual.Contains(expectedSubstring, StringComparison.Ordinal))
        {
            ClassicAssert.Fail(
                message
                    ?? $"Expected string containing \"{expectedSubstring}\" but was \"{actual}\"."
            );
        }
    }

    public static void StartsWith(string expectedPrefix, string? actual, string? message = null)
    {
        if (actual is null || !actual.StartsWith(expectedPrefix, StringComparison.Ordinal))
        {
            ClassicAssert.Fail(
                message
                    ?? $"Expected string starting with \"{expectedPrefix}\" but was \"{actual}\"."
            );
        }
    }
}

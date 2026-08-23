#nullable disable

using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Parser;

namespace XlsxSharp.Extensions;

internal static partial class StringExtensions
{
    [GeneratedRegex(@"((?<!\r)\n|\r\n)")]
    private static partial Regex RegexNewLine { get; }

    public static int CharCount(this string instance, char c) => instance.AsSpan().Count(c);

    public static string RemoveSpecialCharacters(this string str)
    {
        StringBuilder sb = new();
        foreach (char c in str)
        {
            if (char.IsLetterOrDigit(c) || c == '.' || c == '_')
            {
                sb.Append(c);
            }
        }
        return sb.ToString();
    }

    internal static string EscapeSheetName(this string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return sheetName;
        }

        bool needEscape =
            (!char.IsLetter(sheetName[0]) && sheetName[0] != '_')
            || XlsxSharp.XLHelper.IsValidA1Address(sheetName)
            || XlsxSharp.XLHelper.IsValidRCAddress(sheetName)
            || sheetName.Any(c =>
                (char.IsPunctuation(c) && c != '.' && c != '_')
                || char.IsSeparator(c)
                || char.IsControl(c)
                || char.IsSymbol(c)
            );
        if (needEscape)
        {
            return AlwaysEscapeSheetName(sheetName);
        }
        else
        {
            return sheetName;
        }
    }

    internal static string AlwaysEscapeSheetName(this string sheetName) =>
        string.Concat('\'', sheetName.Replace("'", "''"), '\'');

    internal static string GetSheetDefinedName(this string name, string sheet)
    {
        bool shouldEscape = NameUtils.ShouldQuote(sheet.AsSpan());
        string escapedSheetName = shouldEscape ? sheet.AlwaysEscapeSheetName() : sheet;
        return escapedSheetName + '!' + name;
    }

    internal static string FixNewLines(this string value) =>
        value.Contains("\n") ? RegexNewLine.Replace(value, Environment.NewLine) : value;

    internal static bool PreserveSpaces(this string value) =>
        value.StartsWith(' ')
        || value.EndsWith(' ')
        || value.AsSpan().IndexOfAny('\n', '\r', '\t') >= 0;

    internal static string ToCamel(this string value)
    {
        if (value.Length == 0)
        {
            return value;
        }

        return string.Create(
            value.Length,
            value,
            static (span, source) =>
            {
                source.AsSpan().CopyTo(span);
                span[0] = char.ToLower(span[0]);
            }
        );
    }

    internal static string ToProper(this string value)
    {
        if (value.Length == 0)
        {
            return value;
        }

        return string.Create(
            value.Length,
            value,
            static (span, source) =>
            {
                source.AsSpan().CopyTo(span);
                span[0] = char.ToUpper(span[0]);
            }
        );
    }

    internal static string UnescapeSheetName(this string sheetName) =>
        sheetName.Trim('\'').Replace("''", "'");

    /// <summary>
    /// Convert a string (containing code units) into code points.
    /// Surrogate pairs of code units are joined to code points.
    /// </summary>
    /// <param name="text">UTF-16 code units to convert.</param>
    /// <param name="output">Output containing code points. Must always be able to fit whole <paramref name="text"/>.</param>
    /// <returns>Number of code points in the <paramref name="output"/>.</returns>
    internal static int ToCodePoints(this ReadOnlySpan<char> text, Span<int> output)
    {
        int j = 0;
        for (int i = 0; i < text.Length; ++i, ++j)
        {
            if (i + 1 < text.Length && char.IsSurrogatePair(text[i], text[i + 1]))
            {
                output[j] = char.ConvertToUtf32(text[i], text[i + 1]);
                i++;
            }
            else
            {
                output[j] = text[i];
            }
        }

        return j;
    }

    /// <summary>
    /// Is the string a new line of any kind (widnows/unix/mac)?
    /// </summary>
    /// <param name="text">Input text to check for EOL at the beginning.</param>
    /// <param name="length">Length of EOL chars.</param>
    /// <returns>True, if text has EOL at the beginning.</returns>
    internal static bool TrySliceNewLine(this ReadOnlySpan<char> text, out int length)
    {
        if (text.Length >= 2 && text[0] == '\r' && text[1] == '\n')
        {
            length = 2;
            return true;
        }

        if (text.Length >= 1 && (text[0] == '\n' || text[0] == '\r'))
        {
            length = 1;
            return true;
        }

        length = default;
        return false;
    }

    /// <summary>
    /// Convert a magic text to a number, where the first letter is in the highest byte of the number.
    /// </summary>
    internal static uint ToMagicNumber(this string magic)
    {
        if (magic.Length > 4)
        {
            throw new ArgumentException();
        }

        Span<byte> bytes = stackalloc byte[4];
        int written = Encoding.ASCII.GetBytes(magic, bytes);

        uint result = 0;
        for (int i = 0; i < written; i++)
        {
            result = result * 256 + bytes[i];
        }

        return result;
    }

    internal static string TrimFormulaEqual(this string text)
    {
        ReadOnlySpan<char> trimmed = text.AsSpan().Trim();
        if (trimmed.Length >= 1 && trimmed[0] == '=')
        {
            return trimmed[1..].TrimStart().ToString();
        }

        return text;
    }
}

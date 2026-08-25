namespace XlsxSharp.ExcelNumberFormat;

internal static class Token
{
    public static bool IsExponent(ReadOnlySpan<char> token) =>
        token.Equals("e+", StringComparison.OrdinalIgnoreCase)
        || token.Equals("e-", StringComparison.OrdinalIgnoreCase);

    public static bool IsLiteral(ReadOnlySpan<char> token) =>
        (token.Length > 0 && token[0] is '_' or '\\' or '"' or '*')
        || token
            is ","
                or "!"
                or "&"
                or "%"
                or "+"
                or "-"
                or "$"
                or "€"
                or "£"
                or "1"
                or "2"
                or "3"
                or "4"
                or "5"
                or "6"
                or "7"
                or "8"
                or "9"
                or "{"
                or "}"
                or "("
                or ")"
                or " ";

    public static bool IsNumberLiteral(ReadOnlySpan<char> token) =>
        IsPlaceholder(token) || IsLiteral(token) || token is ".";

    public static bool IsPlaceholder(ReadOnlySpan<char> token) =>
        token is "0" or "#" or "?";

    public static bool IsGeneral(ReadOnlySpan<char> token) =>
        token.Equals("general", StringComparison.OrdinalIgnoreCase);

    public static bool IsDatePart(ReadOnlySpan<char> token) =>
        token.StartsWith("y", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("m", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("d", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("s", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("h", StringComparison.OrdinalIgnoreCase)
        || (token.StartsWith("g", StringComparison.OrdinalIgnoreCase) && !IsGeneral(token))
        || token.Equals("am/pm", StringComparison.OrdinalIgnoreCase)
        || token.Equals("a/p", StringComparison.OrdinalIgnoreCase)
        || IsDurationPart(token);

    public static bool IsDurationPart(ReadOnlySpan<char> token) =>
        token.StartsWith("[h", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("[m", StringComparison.OrdinalIgnoreCase)
        || token.StartsWith("[s", StringComparison.OrdinalIgnoreCase);

    public static bool IsDigit09(ReadOnlySpan<char> token) => token is "0" || IsDigit19(token);

    public static bool IsDigit19(ReadOnlySpan<char> token) =>
        token
            is "1"
                or "2"
                or "3"
                or "4"
                or "5"
                or "6"
                or "7"
                or "8"
                or "9";
}

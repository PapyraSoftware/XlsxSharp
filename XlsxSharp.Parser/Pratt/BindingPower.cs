namespace XlsxSharp.Parser.Pratt;

/// <summary>
/// Values of binding power for operators in an expression. Higher number = higher binding power.
/// Precedence of operators is specified by ISO-29500:18.17.2.2. Operators that have the same
/// precedence associate left-to-right.
/// </summary>
internal static class BindingPower
{
    internal const int Comparison = 1;
    internal const int Concat = 2;
    internal const int Addition = 3;
    internal const int Subtraction = 3;
    internal const int Multiplication = 4;
    internal const int Division = 4;
    internal const int Exponentiation = 5;

    /// <summary>
    /// Binding power of the postfix <c>%</c> operator. It is also used as the binding power a
    /// prefix <c>+</c>/<c>-</c> uses to parse its own operand: that operand must stop before a
    /// <c>%</c> or <c>^</c> (both bind looser than a "bare" unary, from the unary's own point of
    /// view - "-2%" is Percent(Minus(2)) and "-2^2" is Pow(Minus(2), 2)), so it has to be at least
    /// this high. Since the comparator used everywhere else is "a following operator applies only
    /// if its binding power is strictly greater than minBp", reusing the same value here excludes
    /// percent (equal, not greater) exactly as intended.
    /// </summary>
    internal const int Percent = 6;
}

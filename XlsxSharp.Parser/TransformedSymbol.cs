using System.Text;

namespace XlsxSharp.Parser;

/// <summary>
/// A symbol that represents a transformed symbol value. Should be used during AST transformation.
/// </summary>
public readonly struct TransformedSymbol
{
    private readonly string _formulaText;

    private readonly SymbolRange _range;

    /// <summary>
    /// The text that replaced symbol or null, if there was no change.
    /// </summary>
    private readonly string? _transformedText;

    private TransformedSymbol(string formulaText, string? transformedText, SymbolRange range) : this()
    {
        this._formulaText = formulaText;
        this._transformedText = transformedText;
        this._range = range;
    }

    /// <summary>
    /// Range of the symbol in original formula.
    /// </summary>
    internal SymbolRange OriginalRange => this._range;

    /// <summary>
    /// Length of the transformed symbol.
    /// </summary>
    internal int Length => this._transformedText?.Length ?? this._range.End - this._range.Start;

    /// <summary>
    /// Create a symbol that is different from what was in the original formula.
    /// </summary>
    /// <param name="formula">Text of whole formula.</param>
    /// <param name="range">Range of the symbol in the formula.</param>
    /// <param name="transformedSymbol">The string of a transformed symbol.</param>
    public static TransformedSymbol ToText(string formula, SymbolRange range, string transformedSymbol)
    {
        return new TransformedSymbol(formula, transformedSymbol, range);
    }

    /// <summary>
    /// Create a new symbol represented by a substring of an original formula. Generally used when
    /// there is no modification of the symbol (i.e. just pass it as it is).
    /// </summary>
    /// <param name="formula">Text of whole formula.</param>
    /// <param name="range">Range of the symbol in the formula.</param>
    public static TransformedSymbol CopyOriginal(string formula, SymbolRange range)
    {
        return new TransformedSymbol(formula, null, range);
    }

    /// <summary>
    /// Get content of the symbol as a span. Doesn't allocate memory.
    /// </summary>
    public ReadOnlySpan<char> AsSpan()
    {
        if (this._transformedText is not null)
        {
            return this._transformedText.AsSpan();
        }

        return this._formulaText.AsSpan(this._range.Start, this._range.End - this._range.Start);
    }

    /// <summary>
    /// Get symbol as a text with extra text at the end.
    /// </summary>
    /// <param name="append">Text to append at the end of the symbol text.</param>
    public string ToString(ReadOnlySpan<char> append)
    {
        if (append.Length == 0 && this._transformedText is not null)
        {
            return this._transformedText;
        }

        ReadOnlySpan<char> text = this.AsSpan();
        return new StringBuilder(text.Length + append.Length)
            .Append(text)
            .Append(append)
            .ToString();
    }

    /// <summary>
    /// Get symbol text representation.
    /// </summary>
    public override string ToString()
    {
        return this.ToString(string.Empty.AsSpan());
    }
}
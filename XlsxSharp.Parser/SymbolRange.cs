namespace XlsxSharp.Parser;

/// <summary>
/// A range of a symbol in formula text.
/// </summary>
public readonly struct SymbolRange
{
    /// <summary>
    /// Create a substring of a symbol.
    /// </summary>
    public SymbolRange(int startIndex, int endIndex)
    {
        this.Start = startIndex;
        this.End = endIndex;
    }

    /// <summary>
    /// Start index of symbol in formula text.
    /// </summary>
    public int Start { get; }

    /// <summary>
    /// End index of symbol in formula text. Can be outside of text bounds, if symbol ends at the
    /// last char of formula.
    /// </summary>
    public int End { get; }

    /// <summary>
    /// Length of a symbol.
    /// </summary>
    public int Length => this.End - this.Start;

    /// <summary>
    /// Get range indexes.
    /// </summary>
    public override string ToString()
    {
        return $"[{this.Start}:{this.End}]";
    }

    internal SymbolRange ExtendRight(SymbolRange rangeToRight)
    {
        if (this.End != rangeToRight.Start)
        {
            throw new InvalidOperationException($"The range end {this.End} doesn't match start of the range to the right {rangeToRight.Start}.");
        }

        return new SymbolRange(this.Start, rangeToRight.End);
    }
}

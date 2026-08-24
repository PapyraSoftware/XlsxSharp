namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// A wildcard is at most 255 chars long text. It can contain <c>*</c> which indicates any number characters (including zero)
/// and <c>?</c> which indicates any single character. If you need to find <c>*</c> or <c>?</c> in a text, prefix them with
/// an escape character <c>~</c>.
/// </summary>
internal readonly struct Wildcard
{
    private readonly string _pattern;
    private readonly bool _unpairedTilda;

    public Wildcard(string pattern)
    {
        this._pattern = pattern;
        int tildes = 0;
        while (
            tildes < this._pattern.Length && this._pattern[this._pattern.Length - tildes - 1] == '~'
        )
        {
            tildes++;
        }

        this._unpairedTilda = tildes % 2 == 1;
    }

    /// <summary>
    /// Search for the wildcard anywhere in the text.
    /// </summary>
    /// <param name="input">Text used to search for a pattern.</param>
    /// <returns>zero-based index of a first character in a text that matches to a pattern or -1, if match wasn't found.</returns>
    public int Search(ReadOnlySpan<char> input)
    {
        ReadOnlySpan<char> pattern = this._pattern.AsSpan();
        if (this._pattern.Length > 255)
        {
            return -1;
        }

        if (this._unpairedTilda)
        {
            pattern = pattern[..^1];
        }

        for (int i = 0; i <= input.Length; i++)
        {
            (bool isMatch, _) = MatchFromStart(pattern, input[i..]);
            if (isMatch)
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Match the pattern against input.
    /// </summary>
    /// <returns>Pattern matches whole input.</returns>
    public bool Matches(ReadOnlySpan<char> input)
    {
        ReadOnlySpan<char> pattern = this._pattern.AsSpan();
        if (this._pattern.Length > 255)
        {
            return false;
        }

        if (this._unpairedTilda)
        {
            pattern = pattern[..^1];
        }

        (bool isMatch, int inputEndIndex) = MatchFromStart(pattern, input);
        return isMatch && inputEndIndex == input.Length;
    }

    /// <summary>
    /// Does the start of an input match the pattern?
    /// </summary>
    private static (bool IsMatch, int InputEndIndex) MatchFromStart(
        ReadOnlySpan<char> pattern,
        ReadOnlySpan<char> input
    )
    {
        int inputIndex = 0;
        int patternIndex = 0;
        int starIndex = -1; // Index of a last processed '*' in the pattern
        int matchIndex = 0; // Input index for last processed '*'. Basically a bookmark for backtracking.

        while (inputIndex < input.Length)
        {
            if (patternIndex < pattern.Length)
            {
                if (pattern[patternIndex] == '?')
                {
                    inputIndex++;
                    patternIndex++;
                    continue;
                }

                if (pattern[patternIndex] == '*')
                {
                    starIndex = patternIndex;
                    matchIndex = inputIndex;
                    patternIndex++;
                    continue;
                }

                if (pattern[patternIndex] == '~' && patternIndex + 1 < pattern.Length)
                {
                    patternIndex++;
                }

                if (
                    char.ToUpperInvariant(pattern[patternIndex])
                    == char.ToUpperInvariant(input[inputIndex])
                )
                {
                    inputIndex++;
                    patternIndex++;
                    continue;
                }
            }

            // Pattern didn't match the input. If there was a previous '*', backtrack and try next position in input.
            if (starIndex != -1)
            {
                matchIndex++;
                inputIndex = matchIndex;
                patternIndex = starIndex + 1;
                continue;
            }

            // No match for pattern char or pattern is complete while input has characters left
            return (patternIndex == pattern.Length, inputIndex);
        }

        // The input has been fully matched. Check for remaining '*' in the pattern
        while (patternIndex < pattern.Length && pattern[patternIndex] == '*')
        {
            patternIndex++;
        }

        return (patternIndex == pattern.Length, inputIndex);
    }
}

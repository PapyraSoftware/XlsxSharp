#nullable disable

// Keep this file CodeMaid organised and cleaned

using System.Diagnostics;
using System.Numerics;

namespace XlsxSharp.Extensions;

internal static class IntegerExtensions
{
    public static bool Between(this int val, int from, int to) => val >= from && val <= to;

    /// <summary>
    /// Get index of highest set bit &lt;= to <paramref name="maximalIndex"/> or -1 if no such bit.
    /// </summary>
    internal static int GetHighestSetBitBelow(this uint value, int maximalIndex)
    {
        Debug.Assert(maximalIndex >= 0 && maximalIndex < 32);

        // Shifting the bit at maximalIndex up to bit 31 discards everything above it, so the
        // remaining leading zeroes say how far below maximalIndex the answer sits. Shifting rather
        // than masking because a mask would need 1u << 32 for maximalIndex 31, which C# wraps to
        // 1u << 0.
        uint shifted = value << (31 - maximalIndex);
        return shifted == 0 ? -1 : maximalIndex - BitOperations.LeadingZeroCount(shifted);
    }

    /// <summary>
    /// Get index of lowest set bit &gt;= to <paramref name="minimalIndex"/> or -1 if no such bit.
    /// </summary>
    internal static int GetLowestSetBitAbove(this uint value, int minimalIndex)
    {
        uint shifted = value >> minimalIndex;
        return shifted == 0 ? -1 : BitOperations.TrailingZeroCount(shifted) + minimalIndex;
    }

    /// <summary>
    /// Get highest set bit index or -1 if no bit is set.
    /// </summary>
    // LeadingZeroCount(0) is 32, so the empty case falls out as -1 without a branch.
    internal static int GetHighestSetBit(this uint value) =>
        31 - BitOperations.LeadingZeroCount(value);
}

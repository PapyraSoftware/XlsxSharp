using System;
using XlsxSharp.Extensions;

namespace XlsxSharp.Tests.Extensions;

/// <summary>
/// The bit helpers are used by the cell slice lookup table, so they are covered indirectly by the
/// slice tests. These pin them down directly: the explicit cases document the contract, and the
/// comparisons against a straightforward reference implementation cover the rest of the input
/// space. Both index parameters are contractually within 0..31, which is what
/// <see cref="IntegerExtensions.GetHighestSetBitBelow"/> asserts, so nothing outside that is tested.
/// </summary>
public class IntegerExtensionsTests
{
    private const int BitCount = 32;

    #region Between

    [Test]
    public void BetweenReturnsTrueForBoundaryValues()
    {
        ClassicAssert.True(5.Between(5, 10));
        ClassicAssert.True(10.Between(5, 10));
    }

    [Test]
    public void BetweenReturnsTrueForValueInsideRange()
    {
        ClassicAssert.True(7.Between(5, 10));
    }

    [Test]
    public void BetweenReturnsFalseForValueOutsideRange()
    {
        ClassicAssert.False(4.Between(5, 10));
        ClassicAssert.False(11.Between(5, 10));
    }

    [Test]
    public void BetweenHandlesSingleValueRange()
    {
        ClassicAssert.True(5.Between(5, 5));
        ClassicAssert.False(4.Between(5, 5));
        ClassicAssert.False(6.Between(5, 5));
    }

    [Test]
    public void BetweenHandlesNegativeAndExtremeValues()
    {
        ClassicAssert.True((-3).Between(-5, -1));
        ClassicAssert.False((-6).Between(-5, -1));
        ClassicAssert.True(int.MinValue.Between(int.MinValue, int.MaxValue));
        ClassicAssert.True(int.MaxValue.Between(int.MinValue, int.MaxValue));
    }

    [Test]
    public void BetweenMatchesReference()
    {
        int[] boundaries = [int.MinValue, -1000, -1, 0, 1, 1000, int.MaxValue];
        foreach (int from in boundaries)
        {
            foreach (int to in boundaries)
            {
                if (from > to)
                {
                    continue;
                }

                foreach (int val in boundaries)
                {
                    ClassicAssert.AreEqual(
                        RefBetween(val, from, to),
                        val.Between(from, to),
                        $"{val} between {from} and {to}"
                    );
                }
            }
        }
    }

    #endregion

    #region GetHighestSetBit

    [Test]
    public void GetHighestSetBitReturnsMinusOneForZero()
    {
        ClassicAssert.AreEqual(-1, 0u.GetHighestSetBit());
    }

    [Test]
    public void GetHighestSetBitFindsEachSingleBit()
    {
        for (int i = 0; i < BitCount; i++)
        {
            ClassicAssert.AreEqual(i, (1u << i).GetHighestSetBit(), $"bit {i}");
        }
    }

    [Test]
    public void GetHighestSetBitIgnoresLowerBits()
    {
        ClassicAssert.AreEqual(31, uint.MaxValue.GetHighestSetBit());
        ClassicAssert.AreEqual(3, 0b1011u.GetHighestSetBit());
        ClassicAssert.AreEqual(31, 0x8000_0001u.GetHighestSetBit());
    }

    [Test]
    public void GetHighestSetBitMatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            ClassicAssert.AreEqual(
                RefHighestSetBit(value),
                value.GetHighestSetBit(),
                $"{value:X8}"
            );
        }
    }

    #endregion

    #region GetHighestSetBitBelow

    [Test]
    public void GetHighestSetBitBelowReturnsMinusOneWhenNoBitAtOrBelowIndex()
    {
        ClassicAssert.AreEqual(-1, 0u.GetHighestSetBitBelow(31));

        // Only bit 5 is set, so anything below it finds nothing.
        ClassicAssert.AreEqual(-1, (1u << 5).GetHighestSetBitBelow(4));
    }

    [Test]
    public void GetHighestSetBitBelowIncludesTheIndexItself()
    {
        ClassicAssert.AreEqual(5, (1u << 5).GetHighestSetBitBelow(5));
        ClassicAssert.AreEqual(0, 1u.GetHighestSetBitBelow(0));
        ClassicAssert.AreEqual(31, (1u << 31).GetHighestSetBitBelow(31));
    }

    [Test]
    public void GetHighestSetBitBelowSkipsBitsAboveTheIndex()
    {
        // Bits 2 and 30 set, capped at 10 -> 2.
        uint value = (1u << 2) | (1u << 30);
        ClassicAssert.AreEqual(2, value.GetHighestSetBitBelow(10));
        ClassicAssert.AreEqual(30, value.GetHighestSetBitBelow(31));
    }

    [Test]
    public void GetHighestSetBitBelowMatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            for (int index = 0; index < BitCount; index++)
            {
                ClassicAssert.AreEqual(
                    RefHighestSetBitBelow(value, index),
                    value.GetHighestSetBitBelow(index),
                    $"{value:X8} below {index}"
                );
            }
        }
    }

    #endregion

    #region GetLowestSetBitAbove

    [Test]
    public void GetLowestSetBitAboveReturnsMinusOneWhenNoBitAtOrAboveIndex()
    {
        ClassicAssert.AreEqual(-1, 0u.GetLowestSetBitAbove(0));

        // Only bit 5 is set, so anything above it finds nothing.
        ClassicAssert.AreEqual(-1, (1u << 5).GetLowestSetBitAbove(6));
    }

    [Test]
    public void GetLowestSetBitAboveIncludesTheIndexItself()
    {
        ClassicAssert.AreEqual(5, (1u << 5).GetLowestSetBitAbove(5));
        ClassicAssert.AreEqual(0, 1u.GetLowestSetBitAbove(0));
        ClassicAssert.AreEqual(31, (1u << 31).GetLowestSetBitAbove(31));
    }

    [Test]
    public void GetLowestSetBitAboveSkipsBitsBelowTheIndex()
    {
        // Bits 2 and 30 set, starting at 10 -> 30.
        uint value = (1u << 2) | (1u << 30);
        ClassicAssert.AreEqual(30, value.GetLowestSetBitAbove(10));
        ClassicAssert.AreEqual(2, value.GetLowestSetBitAbove(0));
    }

    [Test]
    public void GetLowestSetBitAboveMatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            for (int index = 0; index < BitCount; index++)
            {
                ClassicAssert.AreEqual(
                    RefLowestSetBitAbove(value, index),
                    value.GetLowestSetBitAbove(index),
                    $"{value:X8} above {index}"
                );
            }
        }
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Every single bit value, the interesting constants, and a deterministic random sample. The
    /// seed is fixed so that a failure is reproducible.
    /// </summary>
    private static uint[] SampleValues()
    {
        uint[] values = new uint[BitCount + 5 + 512];
        int next = 0;

        for (int i = 0; i < BitCount; i++)
        {
            values[next++] = 1u << i;
        }

        values[next++] = 0u;
        values[next++] = uint.MaxValue;
        values[next++] = 0x8000_0001u;
        values[next++] = 0xAAAA_AAAAu;
        values[next++] = 0x5555_5555u;

        Random random = new(20260822);
        while (next < values.Length)
        {
            values[next++] = (uint)random.NextInt64(uint.MinValue, (long)uint.MaxValue + 1);
        }

        return values;
    }

    private static bool RefBetween(int val, int from, int to) => val >= from && val <= to;

    private static int RefHighestSetBit(uint value) => RefHighestSetBitBelow(value, BitCount - 1);

    private static int RefHighestSetBitBelow(uint value, int maximalIndex)
    {
        for (int i = maximalIndex; i >= 0; i--)
        {
            if ((value & (1u << i)) != 0)
            {
                return i;
            }
        }

        return -1;
    }

    private static int RefLowestSetBitAbove(uint value, int minimalIndex)
    {
        for (int i = minimalIndex; i < BitCount; i++)
        {
            if ((value & (1u << i)) != 0)
            {
                return i;
            }
        }

        return -1;
    }

    #endregion
}

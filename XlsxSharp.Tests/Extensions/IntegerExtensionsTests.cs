using System;
using NUnit.Framework;
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

    #region GetHighestSetBit

    [Test]
    public void GetHighestSetBit_ReturnsMinusOne_ForZero()
    {
        Assert.That(0u.GetHighestSetBit(), Is.EqualTo(-1));
    }

    [Test]
    public void GetHighestSetBit_FindsEachSingleBit()
    {
        for (int i = 0; i < BitCount; i++)
        {
            Assert.That((1u << i).GetHighestSetBit(), Is.EqualTo(i), $"bit {i}");
        }
    }

    [Test]
    public void GetHighestSetBit_IgnoresLowerBits()
    {
        Assert.That(uint.MaxValue.GetHighestSetBit(), Is.EqualTo(31));
        Assert.That(0b1011u.GetHighestSetBit(), Is.EqualTo(3));
        Assert.That(0x8000_0001u.GetHighestSetBit(), Is.EqualTo(31));
    }

    [Test]
    public void GetHighestSetBit_MatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            Assert.That(
                value.GetHighestSetBit(),
                Is.EqualTo(RefHighestSetBit(value)),
                $"{value:X8}"
            );
        }
    }

    #endregion

    #region GetHighestSetBitBelow

    [Test]
    public void GetHighestSetBitBelow_ReturnsMinusOne_WhenNoBitAtOrBelowIndex()
    {
        Assert.That(0u.GetHighestSetBitBelow(31), Is.EqualTo(-1));

        // Only bit 5 is set, so anything below it finds nothing.
        Assert.That((1u << 5).GetHighestSetBitBelow(4), Is.EqualTo(-1));
    }

    [Test]
    public void GetHighestSetBitBelow_IncludesTheIndexItself()
    {
        Assert.That((1u << 5).GetHighestSetBitBelow(5), Is.EqualTo(5));
        Assert.That(1u.GetHighestSetBitBelow(0), Is.EqualTo(0));
        Assert.That((1u << 31).GetHighestSetBitBelow(31), Is.EqualTo(31));
    }

    [Test]
    public void GetHighestSetBitBelow_SkipsBitsAboveTheIndex()
    {
        // Bits 2 and 30 set, capped at 10 -> 2.
        uint value = (1u << 2) | (1u << 30);
        Assert.That(value.GetHighestSetBitBelow(10), Is.EqualTo(2));
        Assert.That(value.GetHighestSetBitBelow(31), Is.EqualTo(30));
    }

    [Test]
    public void GetHighestSetBitBelow_MatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            for (int index = 0; index < BitCount; index++)
            {
                Assert.That(
                    value.GetHighestSetBitBelow(index),
                    Is.EqualTo(RefHighestSetBitBelow(value, index)),
                    $"{value:X8} below {index}"
                );
            }
        }
    }

    #endregion

    #region GetLowestSetBitAbove

    [Test]
    public void GetLowestSetBitAbove_ReturnsMinusOne_WhenNoBitAtOrAboveIndex()
    {
        Assert.That(0u.GetLowestSetBitAbove(0), Is.EqualTo(-1));

        // Only bit 5 is set, so anything above it finds nothing.
        Assert.That((1u << 5).GetLowestSetBitAbove(6), Is.EqualTo(-1));
    }

    [Test]
    public void GetLowestSetBitAbove_IncludesTheIndexItself()
    {
        Assert.That((1u << 5).GetLowestSetBitAbove(5), Is.EqualTo(5));
        Assert.That(1u.GetLowestSetBitAbove(0), Is.EqualTo(0));
        Assert.That((1u << 31).GetLowestSetBitAbove(31), Is.EqualTo(31));
    }

    [Test]
    public void GetLowestSetBitAbove_SkipsBitsBelowTheIndex()
    {
        // Bits 2 and 30 set, starting at 10 -> 30.
        uint value = (1u << 2) | (1u << 30);
        Assert.That(value.GetLowestSetBitAbove(10), Is.EqualTo(30));
        Assert.That(value.GetLowestSetBitAbove(0), Is.EqualTo(2));
    }

    [Test]
    public void GetLowestSetBitAbove_MatchesReference()
    {
        foreach (uint value in SampleValues())
        {
            for (int index = 0; index < BitCount; index++)
            {
                Assert.That(
                    value.GetLowestSetBitAbove(index),
                    Is.EqualTo(RefLowestSetBitAbove(value, index)),
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

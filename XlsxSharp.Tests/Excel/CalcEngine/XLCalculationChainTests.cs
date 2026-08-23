using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel;
using XlsxSharp.Excel.CalcEngine;

namespace XlsxSharp.Tests.Excel.CalcEngine;

public class XlCalculationChainTests
{
    [Test]
    public void EnumeratingEmptyChain()
    {
        XLCalculationChain chain = new();
        CollectionAssert.IsEmpty(GetPoints(chain));
    }

    [Test]
    [Arguments(1)]
    [Arguments(2)]
    [Arguments(3)]
    [Arguments(40)]
    public void EnumeratingWholeChain(int chainLength)
    {
        XLCalculationChain chain = new();
        List<SheetPoint> expectedPoints = [];
        for (int i = 0; i < chainLength; ++i)
        {
            SheetPoint point = new("sheet", new Point(1, i));
            chain.AddLast(point);
            expectedPoints.Add(point);
        }

        CollectionAssert.AreEqual(expectedPoints, GetPoints(chain));
    }

    [Test]
    public void RemoveThrowsOnMissingPoint()
    {
        XLCalculationChain chain = new();

        ClassicAssert.Throws<InvalidOperationException>(() =>
            chain.Remove(new SheetPoint("sheet", new Point(1, 1)))
        );
    }

    [Test]
    public void RemoveLinkFromChain()
    {
        XLCalculationChain chain = new();
        SheetPoint a1 = new("sheet", new Point(1, 1));
        SheetPoint b1 = new("sheet", new Point(1, 2));
        SheetPoint c1 = new("sheet", new Point(1, 3));
        SheetPoint d1 = new("sheet", new Point(1, 4));

        chain.AddLast(a1);
        chain.AddLast(b1);
        chain.AddLast(c1);
        chain.AddLast(d1);

        // Remove point in the middle
        chain.Remove(c1);
        CollectionAssert.AreEqual(new[] { a1, b1, d1 }, GetPoints(chain));

        // Remove last point in the sequence
        chain.Remove(d1);
        CollectionAssert.AreEqual(new[] { a1, b1 }, GetPoints(chain));

        // Remove head
        chain.Remove(a1);
        CollectionAssert.AreEqual(new[] { b1 }, GetPoints(chain));

        // Remove the only remaining
        chain.Remove(b1);
        CollectionAssert.IsEmpty(GetPoints(chain));
    }

    [Test]
    public void AddAfterAddsPoint()
    {
        XLCalculationChain chain = new();
        SheetPoint a1 = new("sheet", new Point(1, 1));
        chain.AddLast(a1);

        // Add as tail for single link chain
        SheetPoint b1 = new("sheet", new Point(1, 2));
        chain.AddAfter(a1, b1, 0);
        CollectionAssert.AreEqual(new[] { a1, b1 }, GetPoints(chain));

        // Add as tail for multi link chain
        SheetPoint c1 = new("sheet", new Point(1, 3));
        chain.AddAfter(b1, c1, 0);
        CollectionAssert.AreEqual(new[] { a1, b1, c1 }, GetPoints(chain));

        // Add somewhere in the middle
        SheetPoint d1 = new("sheet", new Point(1, 4));
        chain.AddAfter(b1, d1, 0);
        CollectionAssert.AreEqual(new[] { a1, b1, d1, c1 }, GetPoints(chain));
    }

    [Test]
    public void MoveToFrontMovesThePointToTheFront()
    {
        XLCalculationChain chain = new();
        SheetPoint a1 = new("sheet", new Point(1, 1));
        chain.AddLast(a1);
        SheetPoint b1 = new("sheet", new Point(1, 2));
        chain.AddLast(b1);
        SheetPoint c1 = new("sheet", new Point(1, 3));
        chain.AddLast(c1);
        SheetPoint d1 = new("sheet", new Point(1, 4));
        chain.AddLast(d1);

        ClassicAssert.True(chain.MoveAhead());
        ClassicAssert.AreEqual(a1, chain.Current);

        // a,b,c,d -> d,a,b,c
        chain.MoveToCurrent(d1);
        ClassicAssert.AreEqual(d1, chain.Current);
        ClassicAssert.AreEqual(new[] { d1, a1, b1, c1 }, GetPoints(chain));

        // d,a,b,c -> b,d,a,c
        chain.MoveToCurrent(b1);
        ClassicAssert.AreEqual(b1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, d1, a1, c1 }, GetPoints(chain));

        ClassicAssert.True(chain.MoveAhead());
        ClassicAssert.AreEqual(d1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, d1, a1, c1 }, GetPoints(chain));

        // d,a,c -> a,d,c
        chain.MoveToCurrent(a1);
        ClassicAssert.AreEqual(a1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, a1, d1, c1 }, GetPoints(chain));

        // Move A1 to front when it's already at front
        chain.MoveToCurrent(a1);
        ClassicAssert.AreEqual(a1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, a1, d1, c1 }, GetPoints(chain));

        // a,d,c -> c,a,d
        chain.MoveToCurrent(c1);
        ClassicAssert.AreEqual(c1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, c1, a1, d1 }, GetPoints(chain));

        ClassicAssert.True(chain.MoveAhead());
        ClassicAssert.AreEqual(a1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, c1, a1, d1 }, GetPoints(chain));

        // a,d -> d,a
        chain.MoveToCurrent(d1);
        ClassicAssert.AreEqual(d1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

        ClassicAssert.True(chain.MoveAhead());
        ClassicAssert.AreEqual(a1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

        // a -> a
        chain.MoveToCurrent(a1);
        ClassicAssert.AreEqual(a1, chain.Current);
        ClassicAssert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));

        ClassicAssert.False(chain.MoveAhead());
        ClassicAssert.AreEqual(new[] { b1, c1, d1, a1 }, GetPoints(chain));
    }

    [Test]
    public void TraversalDetectsCycles()
    {
        XLCalculationChain chain = new();
        // `=C1+B1`
        SheetPoint a1 = new("sheet", new Point(1, 1));
        chain.AddLast(a1);
        // `=A1`
        SheetPoint b1 = new("sheet", new Point(1, 2));
        chain.AddLast(b1);
        // `=A1`
        SheetPoint c1 = new("sheet", new Point(1, 3));
        chain.AddLast(c1);

        // Move to the first link.
        ClassicAssert.True(chain.MoveAhead());

        // Cycle a1, c1, when we first encounter c1, we don't know yet that it's a cycle
        chain.MoveToCurrent(c1);
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));

        // A1 is marked with a position, because they have been at the current
        // C1 hasn't ben pushed back yet, so it keeps 0.
        CollectionAssert.AreEqual(new[] { 0, 1, 0 }, GetPositions(chain));

        // But then we get A1 again, without any other point being marked
        // as done, therefore we are at cycle.
        chain.MoveToCurrent(a1);
        CollectionAssert.AreEqual(new[] { a1, c1, b1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));
        ClassicAssert.True(chain.IsCurrentInCycle);

        // When we encounter C1 again, it's obviously a cycle.
        chain.MoveToCurrent(c1);
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));
        ClassicAssert.True(chain.IsCurrentInCycle);

        // Let's move on and get A1 to the current. Because the C1 has been
        // marked as done, A1 is no longer in cycle.
        chain.MoveAhead();
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));

        // C1 position has been cleared, because it has moved beyond
        // current and A1 is now current.
        CollectionAssert.AreEqual(new[] { 0, 1, 0 }, GetPositions(chain));

        // A1 is no longer in a current, because current position is 2, but last position
        // of A1 was 1 => there has been a processed node in the meantime.
        ClassicAssert.False(chain.IsCurrentInCycle);

        chain.MoveToCurrent(b1);
        CollectionAssert.AreEqual(new[] { c1, b1, a1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 0, 0, 2 }, GetPositions(chain));
        ClassicAssert.False(chain.IsCurrentInCycle);

        chain.MoveToCurrent(a1);
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 0, 2, 2 }, GetPositions(chain));
        ClassicAssert.True(chain.IsCurrentInCycle);

        chain.MoveAhead();
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 0, 0, 2 }, GetPositions(chain));
        ClassicAssert.False(chain.IsCurrentInCycle);

        chain.MoveAhead();
        CollectionAssert.AreEqual(new[] { c1, a1, b1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 0, 0, 0 }, GetPositions(chain));
    }

    [Test]
    public void ResetClearsPositionsAheadOfCurrent()
    {
        XLCalculationChain chain = new();
        SheetPoint a1 = new("sheet", new Point(1, 1));
        chain.AddLast(a1);
        SheetPoint b1 = new("sheet", new Point(1, 2));
        chain.AddLast(b1);
        SheetPoint c1 = new("sheet", new Point(1, 3));
        chain.AddLast(c1);

        ClassicAssert.True(chain.MoveAhead());

        chain.MoveToCurrent(b1);
        chain.MoveToCurrent(a1);
        ClassicAssert.True(chain.IsCurrentInCycle);
        CollectionAssert.AreEqual(new[] { a1, b1, c1 }, GetPoints(chain));
        CollectionAssert.AreEqual(new[] { 1, 1, 0 }, GetPositions(chain));

        chain.Reset();

        CollectionAssert.AreEqual(new[] { 0, 0, 0 }, GetPositions(chain));
    }

    private static IEnumerable<SheetPoint> GetPoints(XLCalculationChain chain) =>
        chain.GetLinks().Select(x => x.Point);

    private static IEnumerable<int> GetPositions(XLCalculationChain chain) =>
        chain.GetLinks().Select(x => x.LastPosition);
}

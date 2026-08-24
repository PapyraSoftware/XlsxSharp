using System.Diagnostics;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// <para>
/// A calculation chain of formulas. Contains all formulas in the workbook.
/// </para>
/// <para>
/// Calculation chain is an ordering of all cells that have value calculated
/// by a formula (note that one formula can determine value of multiple cells,
/// e.g. array). Formulas are calculated in specified order and if currently
/// processed formula needs data from a cell whose value is dirty (i.e. it
/// is determined by a not-yet-calculated formula), the current formula is
/// stopped and the required formula is placed before the current one and starts
/// to be processed. Once it is done, the original formula is starts to be processed
/// again. It might have encounter another not-yet-calculated formula or it
/// will finish and the calculation chain moves to the next one.
/// </para>
/// <para>
/// Chain can be traversed through <see cref="Current"/>, <see cref="MoveAhead"/>,
/// <see cref="MoveToCurrent"/> and <see cref="Reset"/>, but only one traversal
/// can go on at the same time due to shared info about cycle detection.
/// </para>
/// </summary>
internal class XLCalculationChain
{
    /// <summary>
    /// Key to the <see cref="_nodeMap"/> that is the head of the chain.
    /// Null, when chain is empty.
    /// </summary>
    private SheetPoint? _head;

    /// <summary>
    /// Key to the <see cref="_nodeMap"/> that is the tail of the chain.
    /// Null, when chain is empty.
    /// </summary>
    private SheetPoint? _tail;

    /// <summary>
    /// <para>
    /// Doubly circular linked list containing all points with value
    /// calculated by a formula. The chain is "looped", so it doesn't
    /// have to deal with nulls for <see cref="SheetPoint"/>.
    /// </para>
    /// <para>
    /// There is always exactly one loop, no cycles. The formulas might
    /// cause cycles due to dependencies, but that is manifested by
    /// constantly switching the links in a loop.</para>
    /// </summary>
    private readonly Dictionary<SheetPoint, Link> _nodeMap = new();

    private SheetPoint? _current;

    /// <summary>
    /// 1 based position of <see cref="_current"/>, if there is a traversal
    /// in progress (0 otherwise).
    /// </summary>
    private int _currentPosition;

    /// <summary>
    /// The address of a current of the chain.
    /// </summary>
    internal SheetPoint Current => this._current!.Value;

    /// <summary>
    /// Is there a cycle in the chain? Detected when a link has appeared
    /// as a current more than once and the current hasn't moved in the
    /// meantime.
    /// </summary>
    internal bool IsCurrentInCycle { get; private set; }

    /// <summary>
    /// Create a new chain filled with all formulas from the workbook.
    /// </summary>
    internal static XLCalculationChain CreateFrom(XLWorkbook wb)
    {
        XLCalculationChain chain = new();
        foreach (XLWorksheet sheet in wb.WorksheetsInternal)
        {
            FormulaSlice formulaSlice = sheet.Internals.CellsCollection.FormulaSlice;
            using Slice<XLCellFormula>.Enumerator e = formulaSlice.GetForwardEnumerator(Area.Full);
            while (e.MoveNext())
            {
                chain.AddLast(new SheetPoint(sheet.Name, e.Point));
            }
        }

        return chain;
    }

    /// <summary>
    /// Add a new link at the beginning of a chain.
    /// </summary>
    private void AddFirst(SheetPoint point, int lastPosition)
    {
        if (this._head is null || this._tail is null)
        {
            this.Init(point);
            return;
        }

        this.Insert(point, lastPosition, this._tail.Value, this._head.Value);
        this._head = point;
    }

    /// <inheritdoc cref="AddLast(SheetPoint,int)"/>
    internal void AddLast(SheetPoint point) => this.AddLast(point, 0);

    /// <summary>
    /// Add all cells from the area to the end of the chain.
    /// </summary>
    internal void AppendArea(SheetArea area)
    {
        foreach (SheetPoint point in area)
        {
            this.AddLast(point);
        }
    }

    /// <summary>
    /// Append formula at the end of the chain.
    /// </summary>
    private void AddLast(SheetPoint point, int lastPosition)
    {
        if (this._head is null || this._tail is null)
        {
            this.Init(point);
            return;
        }

        this.Insert(point, lastPosition, this._tail.Value, this._head.Value);
        this._tail = point;
    }

    /// <summary>
    /// Initialize empty chain with a single link chain.
    /// </summary>
    private void Init(SheetPoint point)
    {
        Debug.Assert(this._nodeMap.Count == 0 && this._head is null && this._tail is null);
        this._nodeMap.Add(point, new Link(point, point, 0));
        this._head = this._tail = point;
    }

    /// <summary>
    /// Insert a link into the <see cref="_nodeMap"/> between
    /// <paramref name="prev"/> and <paramref name="next"/>.
    /// Don't update head or tail.
    /// </summary>
    private void Insert(SheetPoint point, int lastPosition, SheetPoint prev, SheetPoint next)
    {
        this._nodeMap.Add(point, new Link(prev, next, lastPosition));

        Link prevLink = this._nodeMap[prev];
        this._nodeMap[prev] = new Link(prevLink.Previous, point, prevLink.LastPosition);

        Link nextLink = this._nodeMap[next];
        this._nodeMap[next] = new Link(point, nextLink.Next, nextLink.LastPosition);
    }

    /// <summary>
    /// Add a link for <paramref name="point"/> after the link for
    /// <paramref name="anchor"/>.
    /// </summary>
    /// <param name="anchor">
    /// The anchor point after which will be the new point added.
    /// </param>
    /// <param name="point">Point to add to the chain.</param>
    /// <param name="lastPosition">The last position of the point in the chain.</param>
    internal void AddAfter(SheetPoint anchor, SheetPoint point, int lastPosition)
    {
        Link prevLink = this._nodeMap[anchor];
        SheetPoint next = prevLink.Next;
        this.Insert(point, lastPosition, anchor, next);

        if (anchor == this._tail!.Value)
        {
            this._tail = point;
        }
    }

    /// <summary>
    /// Remove point from the chain.
    /// </summary>
    /// <param name="point">Link to remove.</param>
    /// <returns>Last position of the removed link.</returns>
    /// <exception cref="InvalidOperationException">Point is not a part of the chain.</exception>
    internal int Remove(SheetPoint point)
    {
        if (!this._nodeMap.TryGetValue(point, out Link pointLink))
        {
            throw this.PointNotInChain(point);
        }

        // Point is in the chain and there is exactly one link -> clear all.
        if (this._nodeMap.Count == 1)
        {
            this.Clear();
            return pointLink.LastPosition;
        }

        if (point == this._head!.Value)
        {
            this._head = pointLink.Next;
        }

        if (point == this._tail!.Value)
        {
            this._tail = pointLink.Previous;
        }

        Link prevLink = this._nodeMap[pointLink.Previous];
        Debug.Assert(prevLink.Next == point);
        this._nodeMap[pointLink.Previous] = new Link(
            prevLink.Previous,
            pointLink.Next,
            prevLink.LastPosition
        );

        Link nextLink = this._nodeMap[pointLink.Next];
        Debug.Assert(nextLink.Previous == point);
        this._nodeMap[pointLink.Next] = new Link(
            pointLink.Previous,
            nextLink.Next,
            nextLink.LastPosition
        );

        this._nodeMap.Remove(point);
        return pointLink.LastPosition;
    }

    /// <summary>
    /// Clear whole chain.
    /// </summary>
    internal void Clear()
    {
        this._nodeMap.Clear();
        this._head = null;
        this._tail = null;
    }

    /// <summary>
    /// Enumerate all links in the chain.
    /// </summary>
    internal IEnumerable<(SheetPoint Point, int LastPosition)> GetLinks()
    {
        if (this._head is null)
        {
            yield break;
        }

        SheetPoint current = this._head.Value;
        do
        {
            Link link = this._nodeMap[current];
            yield return new ValueTuple<SheetPoint, int>(current, link.LastPosition);
            current = link.Next;
        } while (current != this._head.Value);
    }

    internal void Reset()
    {
        if (this._current is null)
        {
            return;
        }

        SheetPoint point = this._current.Value;
        Link link = this._nodeMap[point];
        while (link.LastPosition != 0)
        {
            this._nodeMap[point] = new Link(link.Previous, link.Next, 0);
            point = link.Next;
            link = this._nodeMap[point];
        }

        this._current = null;
        this._currentPosition = 0;
    }

    /// <summary>
    /// Mark current link as complete and move ahead to the next link.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the enumerator moved ahead, <c>false</c> if
    /// there are no more links and chain has looped completely.
    /// </returns>
    internal bool MoveAhead()
    {
        // First move
        if (this._current is null)
        {
            bool isChainEmpty = this._head is null;
            if (isChainEmpty)
            {
                return false;
            }

            this._current = this._head;
            this._currentPosition = 1;
            return true;
        }

        // Subsequent move
        SheetPoint currentPoint = this._current.Value;
        if (!this._nodeMap.TryGetValue(currentPoint, out Link currentLink))
        {
            throw this.PointNotInChain(currentPoint);
        }

        // Clear up the last position, the current point is being moved to done
        // and clearing will ensure next traversal won't be affected.
        if (currentLink.LastPosition != 0)
        {
            this._nodeMap[currentPoint] = new Link(currentLink.Previous, currentLink.Next, 0);
        }

        SheetPoint nextPoint = currentLink.Next;
        Debug.Assert(this._nodeMap[nextPoint].Previous == currentPoint);
        if (nextPoint == this._head!.Value)
        {
            // Whole chain has been calculated.
            return false;
        }

        // Since we moved, the new last position is greater than all others
        // and thus can't be in the cycle.
        this.IsCurrentInCycle = false;
        this._current = nextPoint;
        this._currentPosition++;
        return true;
    }

    /// <summary>
    /// Move the <paramref name="pointToMove"/> before the current point
    /// as the new current to be calculated.
    /// </summary>
    /// <param name="pointToMove">
    /// The point of a chain to moved to the current. Should always be in
    /// the chain after the current.
    /// </param>
    internal void MoveToCurrent(SheetPoint pointToMove)
    {
        if (this._current is null)
        {
            throw new InvalidOperationException("Enumerator not at a link.");
        }

        SheetPoint currentPoint = this._current.Value;

        // If we are not moving anything, adding and removing doesn't
        // change chain, plus we avoid problems with moving in a
        // single/double link chain.
        if (currentPoint == pointToMove)
        {
            // But it basically means that currentPoint depends on pointToMove
            // thus cell depends on itself and that is a cycle.
            this.IsCurrentInCycle = true;
            return;
        }

        // If head is also current, moving before the current means moving before head
        int pointToMoveLastPosition = this.Remove(pointToMove);
        if (this._head == currentPoint)
        {
            this.AddFirst(pointToMove, pointToMoveLastPosition);
        }
        else
        {
            // Current is not a head = move a link after prev of current.
            SheetPoint anchor = this._nodeMap[currentPoint].Previous;
            this.AddAfter(anchor, pointToMove, pointToMoveLastPosition);
        }

        Link shiftedLink = this._nodeMap[currentPoint];
        this._nodeMap[currentPoint] = new Link(
            shiftedLink.Previous,
            shiftedLink.Next,
            this._currentPosition
        );

        this.IsCurrentInCycle = this._currentPosition == pointToMoveLastPosition;
        this._current = pointToMove;
    }

    private InvalidOperationException PointNotInChain(SheetPoint point)
    {
        InvalidOperationException exception = new($"Book point {point} is not in the chain.");
        exception.Data.Add(
            "Chain",
            string.Join(
                ", ",
                this._nodeMap.Select(n => $"{n.Key}(prev:{n.Value.Previous},next:{n.Value.Next})")
            )
        );
        return exception;
    }

    private readonly struct Link
    {
        internal readonly SheetPoint Previous;

        internal readonly SheetPoint Next;

        /// <summary>
        /// <para>
        /// What was the 1-based position of the link in the chain the last
        /// time the link has been current. Only used when link is pushed
        /// to the back, otherwise it's <c>0</c>.
        /// </para>
        /// <para>
        /// The last position of a link is only updated when
        /// <list type="bullet">
        /// <item>
        /// Link is moved from current to the back - that means link
        /// will be moved to current again at some point in the future
        /// and if chain hasn't processed even one link in the meantime,
        /// there is a cycle.
        /// </item>
        /// <item>
        /// Link is marked as done and current moves past it. The last
        /// position should be cleared as not to confuse next traversal.
        /// </item>
        /// <item>
        /// Chain traversal is reset - links in front of current may still
        /// have set their last position, because other links have been
        /// moved to the current as a supporting links.
        /// </item>
        /// </list>
        /// </para>
        /// </summary>
        /// <remarks>Used for cycle detection.</remarks>
        internal readonly int LastPosition;

        public Link(SheetPoint previous, SheetPoint next, int lastPosition)
        {
            this.Previous = previous;
            this.Next = next;
            this.LastPosition = lastPosition;
        }
    }
}

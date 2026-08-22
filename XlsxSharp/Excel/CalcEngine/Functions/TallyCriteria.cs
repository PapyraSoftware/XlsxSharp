using System;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel.CalcEngine.Functions;

/// <summary>
/// Tally for <c>{SUM,COUNT,AVERAGE}IF/S</c> and database function. The created tally must contain
/// all selection areas and associated criteria. The main <see cref="Tally{T}"/> function is then
/// called with values that will be tallied, based on the areas+criteria in the tally object.
/// </summary>
internal class TallyCriteria : ITally
{
    /// <summary>
    /// A collection of areas that are tested and if all satisfy the criteria, corresponding values
    /// in the tally areas are tallied.
    /// </summary>
    private readonly List<(XLRangeAddress Area, Criteria Criteria)> _criteriaRanges = [];

    /// <summary>
    /// A method to convert a value in the tally area to a number. If scalar value shouldn't be tallied, return null.
    /// </summary>
    private readonly Func<ScalarValue, double?> _toNumber;

    internal TallyCriteria()
        : this(static cellValue => cellValue.TryPickNumber(out double number) ? number : null) { }

    internal TallyCriteria(Func<ScalarValue, double?> toNumber) => this._toNumber = toNumber;

    /// <summary>
    /// Add criteria to the tally that limit which values should be tallied.
    /// </summary>
    internal void Add(XLRangeAddress area, Criteria criteria) =>
        this._criteriaRanges.Add((area, criteria));

    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        // All criteria functions permit only area reference arguments. Excel ensures this
        // invariant by grammar, we just check the the argument value.
        List<XLRangeAddress> talliedAreas = new(args.Length);
        foreach (AnyValue arg in args)
        {
            ctx.ThrowIfCancelled();
            if (!arg.TryPickArea(out XLRangeAddress tallyArea, out XLError error))
            {
                return error;
            }

            talliedAreas.Add(tallyArea);
        }

        // For each selection area and its criteria, get list of points that satisfy the criteria.
        List<(Point Origin, IEnumerable<Point> Enumerable)> criteriaPoints = [];
        foreach ((XLRangeAddress area, Criteria criteria) in this._criteriaRanges)
        {
            // This is a lazy IEnumerable, it's not yet evaluated.
            IEnumerable<Point> areaCriteriaPoints = ctx.GetCriteriaPoints(area, criteria);
            Point origin = Area.FromRangeAddress(area).FirstPoint;
            criteriaPoints.Add((origin, areaCriteriaPoints));
        }

        // Get list of points that satisfy all criteria
        IEnumerable<XLSheetOffset> talliedCoordinates = GetCombinedCoordinates(criteriaPoints);

        T state = initialState;
        foreach ((int rowOfs, int colOfs) in talliedCoordinates)
        {
            foreach (XLRangeAddress area in talliedAreas)
            {
                ctx.ThrowIfCancelled();
                XLAddress origin = area.FirstAddress;
                Point shifted = new(origin.RowNumber + rowOfs, origin.ColumnNumber + colOfs);
                ScalarValue cellValue = ctx.GetCellValue(
                    area.Worksheet,
                    shifted.Row,
                    shifted.Column
                );
                double? number = this._toNumber(cellValue);
                if (number is not null)
                {
                    state = state.Tally(number.Value);
                }
            }
        }

        return state;
    }

    private static IEnumerable<XLSheetOffset> GetCombinedCoordinates(
        List<(Point Origin, IEnumerable<Point> Enumerable)> enumerables
    )
    {
        List<IEnumerator<Point>> enumerators =
        [
            .. enumerables.Select(e => e.Enumerable.GetEnumerator()),
        ];
        try
        {
            // Move to the first element
            foreach (IEnumerator<Point> enumerator in enumerators)
            {
                if (!enumerator.MoveNext())
                {
                    yield break;
                }
            }

            // Until all elements are processed.
            while (true)
            {
                // Do all enumerators have same offset?
                bool allSame = true;
                XLSheetOffset minOfs = GetOffset(0);
                for (int i = 1; i < enumerables.Count; ++i)
                {
                    XLSheetOffset currentOfs = GetOffset(i);
                    int comparison = currentOfs.CompareTo(minOfs);
                    if (minOfs != currentOfs)
                    {
                        allSame = false;
                    }

                    if (comparison < 0)
                    {
                        minOfs = currentOfs;
                    }
                }

                // If all offsets are same, that means all criteria are
                // satisfied for same offset.
                if (allSame)
                {
                    yield return minOfs;
                }

                // Move all enumerators that point at the minimum offset
                // to the next element.
                for (int i = 0; i < enumerables.Count; ++i)
                {
                    XLSheetOffset currentOfs = GetOffset(i);
                    if (currentOfs.CompareTo(minOfs) <= 0)
                    {
                        if (!enumerators[i].MoveNext())
                        {
                            yield break;
                        }
                    }
                }
            }
        }
        finally
        {
            foreach (IEnumerator<Point> enumerator in enumerators)
            {
                enumerator.Dispose();
            }
        }

        XLSheetOffset GetOffset(int i)
        {
            Point origin = enumerables[i].Origin;
            Point point = enumerators[i].Current;
            return point - origin;
        }
    }
}

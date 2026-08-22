using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XlsxSharp.Excel;

/// <summary>
/// Engine for ranges consolidation. Supports IXLRanges including ranges from either one or multiple worksheets.
/// </summary>
internal class XLRangeConsolidationEngine
{
    private readonly XLWorkbook _workbook;
    private readonly XLRanges _allRanges;

    public XLRangeConsolidationEngine(XLWorkbook workbook, XLRanges ranges)
    {
        this._workbook = workbook;
        this._allRanges = ranges ?? throw new ArgumentNullException(nameof(ranges));
    }

    public XLRanges Consolidate()
    {
        if (this._allRanges.Count == 0)
        {
            return this._allRanges;
        }

        IOrderedEnumerable<XLWorksheet> worksheets = this
            ._allRanges.Select<XLRange, XLWorksheet>(r => r.Worksheet)
            .Distinct()
            .OrderBy(ws => ws.Position);

        XLRanges retVal = new(this._workbook);
        foreach (XLWorksheet ws in worksheets)
        {
            XLAreaList areaList = new([
                .. this._allRanges.Where<XLRange>(r => r.Worksheet == ws).Select(r => r.SheetRange),
            ]);
            XLRangeConsolidationMatrix matrix = new(areaList);
            IEnumerable<Area> consRanges = matrix.GetConsolidatedRanges();
            foreach (Area consArea in consRanges)
            {
                retVal.Add(ws.Range(consArea));
            }
        }

        return retVal;
    }

    internal static XLAreaList Consolidate(XLAreaList areas)
    {
        if (areas.Count == 0)
        {
            return areas;
        }

        XLRangeConsolidationMatrix matrix = new(areas);
        List<Area> consRanges = [.. matrix.GetConsolidatedRanges()];
        return new XLAreaList(consRanges);
    }

    /// <summary>
    /// Class representing the area covering ranges to be consolidated as a set of bit matrices. Does all the dirty job
    /// of ranges consolidation.
    /// </summary>
    private class XLRangeConsolidationMatrix
    {
        private readonly Dictionary<int, BitArray> _bitMatrix;
        private readonly int _minColumn;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="areas">Areas to be consolidated.</param>
        internal XLRangeConsolidationMatrix(XLAreaList areas)
        {
            (this._bitMatrix, this._minColumn) = PrepareBitMatrix(areas);
            this.FillBitMatrix(areas);
        }

        /// <summary>
        /// Get consolidated ranges equivalent to the input ones.
        /// </summary>
        public IEnumerable<Area> GetConsolidatedRanges()
        {
            int[] rowNumbers = [.. this._bitMatrix.Keys.OrderBy(k => k)];
            for (int i = 0; i < rowNumbers.Length; i++)
            {
                int startRow = rowNumbers[i];
                IEnumerable<Tuple<int, int>> startings = GetRangesBoundariesStartingByRow(
                    this._bitMatrix[startRow]
                );

                foreach (Tuple<int, int> starting in startings)
                {
                    int j = i + 1;
                    while (
                        j < rowNumbers.Length
                        && RowIncludesRange(this._bitMatrix[rowNumbers[j]], starting)
                    )
                    {
                        j++;
                    }

                    int endRow = rowNumbers[j - 1];
                    int startColumn = starting.Item1 + this._minColumn - 1;
                    int endColumn = starting.Item2 + this._minColumn - 1;

                    yield return new Area(startRow, startColumn, endRow, endColumn);

                    while (j > i)
                    {
                        ClearRangeInRow(this._bitMatrix[rowNumbers[j - 1]], starting);
                        j--;
                    }
                }
            }
        }

        private void AddToBitMatrix(Area area)
        {
            IEnumerable<int> rows = this._bitMatrix.Keys.Where(k =>
                k >= area.TopRow && k <= area.BottomRow
            );

            int minIndex = area.LeftColumn - this._minColumn + 1;
            int maxIndex = area.RightColumn - this._minColumn + 1;

            foreach (int rowNum in rows)
            {
                for (int i = minIndex; i <= maxIndex; i++)
                {
                    this._bitMatrix[rowNum][i] = true;
                }
            }
        }

        private static void ClearRangeInRow(BitArray rowArray, Tuple<int, int> rangeBoundaries)
        {
            for (int i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
            {
                rowArray[i] = false;
            }
        }

        private void FillBitMatrix(IEnumerable<Area> areas)
        {
            foreach (Area area in areas)
            {
                this.AddToBitMatrix(area);
            }

            System.Diagnostics.Debug.Assert(
                this._bitMatrix.Values.All(r => r[0] == false && r[r.Length - 1] == false)
            );
        }

        private static IEnumerable<Tuple<int, int>> GetRangesBoundariesStartingByRow(
            BitArray rowArray
        )
        {
            int startIdx = 0;
            for (int i = 1; i < rowArray.Length - 1; i++)
            {
                if (!rowArray[i - 1] && rowArray[i])
                {
                    startIdx = i;
                }

                if (rowArray[i] && !rowArray[i + 1])
                {
                    yield return new Tuple<int, int>(startIdx, i);
                }
            }
        }

        private static (Dictionary<int, BitArray> BitMatrix, int MinColumn) PrepareBitMatrix(
            XLAreaList areas
        )
        {
            int minColumn = XlsxSharp.XLHelper.MaxColumnNumber + 1;
            int maxColumn = 0;
            foreach (Area area in areas)
            {
                minColumn = (minColumn <= area.LeftColumn) ? minColumn : area.LeftColumn;
                maxColumn = (maxColumn >= area.RightColumn) ? maxColumn : area.RightColumn;
            }

            int bitMaskSize = maxColumn - minColumn + 3;
            Dictionary<int, BitArray> bitMatrix = new();
            foreach (Area area in areas)
            {
                AddRowBitmask(bitMatrix, area.TopRow, bitMaskSize);
                AddRowBitmask(bitMatrix, area.BottomRow, bitMaskSize);
                AddRowBitmask(bitMatrix, area.BottomRow + 1, bitMaskSize);
            }

            return (bitMatrix, minColumn);

            static void AddRowBitmask(
                Dictionary<int, BitArray> bitMatrix,
                int rowNum,
                int bitMaskSize
            )
            {
                if (!bitMatrix.ContainsKey(rowNum))
                {
                    bitMatrix.Add(rowNum, new BitArray(bitMaskSize, false));
                }
            }
        }

        private static bool RowIncludesRange(BitArray rowArray, Tuple<int, int> rangeBoundaries)
        {
            for (int i = rangeBoundaries.Item1; i <= rangeBoundaries.Item2; i++)
            {
                if (!rowArray[i])
                {
                    return false;
                }
            }

            return true;
        }
    }
}

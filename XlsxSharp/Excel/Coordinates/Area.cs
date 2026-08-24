using System.Collections;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;

namespace XlsxSharp.Excel;

/// <summary>
/// A representation of a <c>ST_Ref</c>, i.e. an area in a sheet (no reference to the sheet).
/// </summary>
internal readonly struct Area : IEquatable<Area>, IEnumerable<Point>
{
    internal Area(Point point)
        : this(point, point) { }

    internal Area(Point firstPoint, Point lastPoint)
    {
        this.FirstPoint = firstPoint;
        this.LastPoint = lastPoint;
    }

    public Area(int rowStart, int columnStart, int rowEnd, int columnEnd)
        : this(new Point(rowStart, columnStart), new Point(rowEnd, columnEnd)) { }

    /// <summary>
    /// A range that covers whole worksheet.
    /// </summary>
    public static readonly Area Full = new(
        new Point(XlsxSharp.XLHelper.MinRowNumber, XlsxSharp.XLHelper.MinColumnNumber),
        new Point(XlsxSharp.XLHelper.MaxRowNumber, XlsxSharp.XLHelper.MaxColumnNumber)
    );

    /// <summary>
    /// Top-left point of the sheet range.
    /// </summary>
    public readonly Point FirstPoint;

    /// <summary>
    /// Bottom-right point of the sheet range.
    /// </summary>
    public readonly Point LastPoint;

    public int Width => this.LastPoint.Column - this.FirstPoint.Column + 1;

    public int Height => this.LastPoint.Row - this.FirstPoint.Row + 1;

    /// <summary>
    /// The left column number of the range. From 1 to <see cref="XlsxSharp.XLHelper.MaxColumnNumber"/>.
    /// </summary>
    public int LeftColumn => this.FirstPoint.Column;

    /// <summary>
    /// The right column number of the range. From 1 to <see cref="XlsxSharp.XLHelper.MaxColumnNumber"/>.
    /// Greater or equal to <see cref="LeftColumn"/>.
    /// </summary>
    public int RightColumn => this.LastPoint.Column;

    /// <summary>
    /// The top row number of the range. From 1 to <see cref="XlsxSharp.XLHelper.MaxRowNumber"/>.
    /// </summary>
    public int TopRow => this.FirstPoint.Row;

    /// <summary>
    /// The bottom row number of the range. From 1 to <see cref="XlsxSharp.XLHelper.MaxRowNumber"/>.
    /// Greater or equal to <see cref="TopRow"/>.
    /// </summary>
    public int BottomRow => this.LastPoint.Row;

    /// <summary>
    /// Does area span from first to last column?
    /// </summary>
    internal bool HasFullRowWidth =>
        this.LeftColumn == XlsxSharp.XLHelper.MinColumnNumber
        && this.RightColumn == XlsxSharp.XLHelper.MaxColumnNumber;

    /// <summary>
    /// Does area span from first to last row?
    /// </summary>
    internal bool HasFullColumnHeight =>
        this.TopRow == XlsxSharp.XLHelper.MinRowNumber
        && this.BottomRow == XlsxSharp.XLHelper.MaxRowNumber;

    public override bool Equals(object? obj) => obj is Area range && this.Equals(range);

    public bool Equals(Area other) =>
        this.FirstPoint.Equals(other.FirstPoint) && this.LastPoint.Equals(other.LastPoint);

    public override int GetHashCode() =>
        this.FirstPoint.GetHashCode() ^ this.LastPoint.GetHashCode();

    public static bool operator ==(Area left, Area right) => left.Equals(right);

    public static bool operator !=(Area left, Area right) => !(left == right);

    /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
    public static Area Parse(string input) => Parse(input.AsSpan());

    /// <summary>
    /// Parse point per type <c>ST_Ref</c> from
    /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/e7f22870-88a1-4c06-8e5f-d035b1179c50">2.1.1119 Part 4 Section 3.18.64, ST_Ref (Cell Range Reference)</a>
    /// </summary>
    /// <remarks>Can be one cell reference (A1) or two separated by a colon (A1:B2). First reference is always in top left corner</remarks>
    /// <param name="input">Input text</param>
    /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
    public static Area Parse(ReadOnlySpan<char> input)
    {
        if (!TryParse(input, out Area area))
        {
            throw new FormatException(
                $"Area reference doesn't have correct format: '{input.ToString()}'."
            );
        }

        return area;
    }

    /// <summary>
    /// Try to parse area. Doesn't accept any extra whitespace anywhere in the input. Letters
    /// must be upper case. Area can specify one corner (<c>A1</c>) or both corners (<c>A1:B3</c>).
    /// </summary>
    public static bool TryParse(ReadOnlySpan<char> input, out Area area)
    {
        int separatorIndex = input.IndexOf(':');
        if (separatorIndex == -1)
        {
            if (!Point.TryParse(input, out Point sheetPoint))
            {
                area = default;
                return false;
            }

            area = new Area(sheetPoint, sheetPoint);
            return true;
        }

        if (
            !Point.TryParse(input[..separatorIndex], out Point first)
            || !Point.TryParse(input[(separatorIndex + 1)..], out Point second)
            || first.Column > second.Column
            || first.Row > second.Row
        )
        {
            area = default;
            return false;
        }

        area = new Area(first, second);
        return true;
    }

    /// <summary>
    /// Write the sheet range to the span. If range has only one cell, write only the cell.
    /// </summary>
    /// <param name="output">Must be at least 21 chars long.</param>
    /// <returns>Number of written characters.</returns>
    public int Format(Span<char> output)
    {
        if (this.FirstPoint == this.LastPoint)
        {
            return this.FirstPoint.Format(output);
        }

        int firstPointLen = this.FirstPoint.Format(output);
        output[firstPointLen] = ':';
        int lastPointLen = this.LastPoint.Format(output.Slice(firstPointLen + 1));
        return firstPointLen + 1 + lastPointLen;
    }

    public override string ToString()
    {
        Span<char> text = stackalloc char[21];
        int len = this.Format(text);
        return text.Slice(0, len).ToString();
    }

    /// <summary>
    /// Return a range that contains all cells below the current range.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
    internal Area BelowRange() => this.BelowRange(XlsxSharp.XLHelper.MaxRowNumber);

    /// <summary>
    /// Get a range below the current one <paramref name="rows"/> rows.
    /// If there isn't enough rows, use as many as possible.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the bottom border of the sheet.</exception>
    internal Area BelowRange(int rows)
    {
        if (this.LastPoint.Row >= XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new InvalidOperationException("No cells below.");
        }

        rows = Math.Min(rows, XlsxSharp.XLHelper.MaxRowNumber - this.LastPoint.Row);
        return new Area(
            new Point(this.LastPoint.Row + 1, this.FirstPoint.Column),
            new Point(this.LastPoint.Row + rows, this.LastPoint.Column)
        );
    }

    /// <summary>
    /// Return a range that contains all cells to the right of the range.
    /// </summary>
    /// <exception cref="InvalidOperationException">The range touches the right border of the sheet.</exception>
    internal Area RightRange()
    {
        if (this.LastPoint.Column == XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new InvalidOperationException("No cells to the left.");
        }

        return new Area(
            new Point(this.FirstPoint.Row, this.LastPoint.Column + 1),
            new Point(this.LastPoint.Row, XlsxSharp.XLHelper.MaxColumnNumber)
        );
    }

    /// <summary>
    /// Return a range that contains additional number of rows below.
    /// </summary>
    internal Area ExtendBelow(int rows)
    {
        Debug.Assert(rows >= 0);
        int row = Math.Min(this.LastPoint.Row + rows, XlsxSharp.XLHelper.MaxRowNumber);
        return new Area(this.FirstPoint, new Point(row, this.LastPoint.Column));
    }

    /// <summary>
    /// Return a range that contains additional number of columns to the right.
    /// </summary>
    internal Area ExtendRight(int columns)
    {
        Debug.Assert(columns >= 0);
        int column = Math.Min(this.LastPoint.Column + columns, XlsxSharp.XLHelper.MaxColumnNumber);
        return new Area(this.FirstPoint, new Point(this.LastPoint.Row, column));
    }

    internal static Area FromRangeAddress<T>(T address)
        where T : IXLRangeAddress
    {
        Point firstPoint = Point.FromAddress(address.FirstAddress);
        Point lastPoint = Point.FromAddress(address.LastAddress);
        if (firstPoint.Row > lastPoint.Row || firstPoint.Column > lastPoint.Column)
        {
            return new Area(lastPoint, firstPoint);
        }

        return new Area(firstPoint, lastPoint);
    }

    public bool Contains(Point point) =>
        point.Row >= this.FirstPoint.Row
        && point.Row <= this.LastPoint.Row
        && point.Column >= this.FirstPoint.Column
        && point.Column <= this.LastPoint.Column;

    internal bool Covers(Area otherArea) =>
        this.LeftColumn <= otherArea.LeftColumn
        && this.TopRow <= otherArea.TopRow
        && this.RightColumn >= otherArea.RightColumn
        && this.BottomRow >= otherArea.BottomRow;

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the bottom row up.
    /// </summary>
    /// <param name="rows">How many rows to take, must be at least one.</param>
    public Area SliceFromBottom(int rows)
    {
        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException();
        }

        return new Area(
            new Point(this.BottomRow - rows + 1, this.FirstPoint.Column),
            this.LastPoint
        );
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the top row down.
    /// </summary>
    /// <param name="rows">How many rows to take, must be at least one.</param>
    public Area SliceFromTop(int rows)
    {
        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException();
        }

        return new Area(this.FirstPoint, new Point(this.TopRow + rows - 1, this.LastPoint.Column));
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the left column to the right.
    /// </summary>
    /// <param name="columns">How many columns to take, must be at least one.</param>
    public Area SliceFromLeft(int columns)
    {
        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException();
        }

        return new Area(
            this.FirstPoint,
            new Point(this.LastPoint.Row, this.LeftColumn + columns - 1)
        );
    }

    /// <summary>
    /// Create a new range from this one by taking a number of rows from the bottom row up.
    /// </summary>
    /// <param name="columns">How many columns to take, must be at least one.</param>
    public Area SliceFromRight(int columns)
    {
        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException();
        }

        return new Area(
            new Point(this.FirstPoint.Row, this.RightColumn - columns + 1),
            this.LastPoint
        );
    }

    /// <summary>
    /// Create a new sheet range that is a result of range operator (<c>:</c>)
    /// of this sheet range and <paramref name="otherRange"/>
    /// </summary>
    /// <param name="otherRange">The other range.</param>
    /// <returns>A range that contains both this range and <paramref name="otherRange"/>.</returns>
    public Area Range(Area otherRange)
    {
        int topRow = Math.Min(this.TopRow, otherRange.TopRow);
        int leftColumn = Math.Min(this.LeftColumn, otherRange.LeftColumn);
        int bottomRow = Math.Max(this.BottomRow, otherRange.BottomRow);
        int rightColumn = Math.Max(this.RightColumn, otherRange.RightColumn);
        return new Area(topRow, leftColumn, bottomRow, rightColumn);
    }

    /// <summary>
    /// Does this range intersects with <paramref name="other"/>.
    /// </summary>
    /// <returns><c>true</c> if intersects, <c>false</c> otherwise.</returns>
    internal bool Intersects(Area other) => this.Intersect(other) is not null;

    /// <summary>
    /// Do an intersection between this range and other range.
    /// </summary>
    /// <param name="other">Other range.</param>
    /// <returns>The intersection range if it exists and is non-empty or null, if intersection doesn't exist.</returns>
    internal Area? Intersect(Area other)
    {
        int leftColumn = Math.Max(this.LeftColumn, other.LeftColumn);
        int rightColumn = Math.Min(this.RightColumn, other.RightColumn);
        int topRow = Math.Max(this.TopRow, other.TopRow);
        int bottomRow = Math.Min(this.BottomRow, other.BottomRow);

        if (bottomRow < topRow || rightColumn < leftColumn)
        {
            return null;
        }

        return new Area(topRow, leftColumn, bottomRow, rightColumn);
    }

    /// <summary>
    /// Does this range overlaps the <paramref name="otherRange"/>?
    /// </summary>
    internal bool Overlaps(Area otherRange) =>
        this.TopRow <= otherRange.TopRow
        && this.RightColumn >= otherRange.RightColumn
        && this.BottomRow >= otherRange.BottomRow
        && this.LeftColumn <= otherRange.LeftColumn;

    /// <summary>
    /// Does range cover all rows, from top row to bottom row of a sheet.
    /// </summary>
    internal bool IsEntireColumn() =>
        this.TopRow == 1 && this.BottomRow == XlsxSharp.XLHelper.MaxRowNumber;

    /// <summary>
    /// Does range cover all columns, from first to last column of a sheet.
    /// </summary>
    public bool IsEntireRow() =>
        this.LeftColumn == 1 && this.RightColumn == XlsxSharp.XLHelper.MaxColumnNumber;

    /// <summary>
    /// Return a new range that has the same size as the current one,
    /// </summary>
    /// <param name="topLeftCorner">New top left coordinate of returned range.</param>
    /// <returns>New range.</returns>
    internal Area At(Point topLeftCorner)
    {
        Point bottomRightCorner = topLeftCorner
            .ShiftColumn(this.Width - 1)
            .ShiftRow(this.Height - 1);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    /// <summary>
    /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/> and in horizontal direction by <paramref name="columnShift"/>.
    /// </summary>
    /// <param name="rowShift">By how much to shift the range.</param>
    /// <param name="columnShift">By how many columns to shift the range.</param>
    /// <returns>Newly created area.</returns>
    internal Area? ShiftAndClip(int rowShift, int columnShift)
    {
        if (this.ShiftRowsAndClip(rowShift) is not { } rowShifted)
        {
            return null;
        }

        if (rowShifted.ShiftColumnsAndClip(columnShift) is not { } rowAndColumnShifted)
        {
            return null;
        }

        return rowAndColumnShifted;
    }

    /// <summary>
    /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/>.
    /// </summary>
    /// <param name="rowShift">By how much to shift the range, positive - downwards, negative - upwards.</param>
    /// <returns>Newly created area.</returns>
    internal Area ShiftRows(int rowShift)
    {
        Point topLeftCorner = this.FirstPoint.ShiftRow(rowShift);
        Point bottomRightCorner = this.LastPoint.ShiftRow(rowShift);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    /// <summary>
    /// Return a new range that has been shifted in vertical direction by <paramref name="rowShift"/>.
    /// If the shifted area is out of sheet bounds, clip part that is out.
    /// </summary>
    /// <param name="rowShift">How many rows to shift.</param>
    /// <returns>Shifted clipped area or <c>null</c> if area was shifted completely out of a sheet.</returns>
    internal Area? ShiftRowsAndClip(int rowShift)
    {
        int shiftedTop = this.TopRow + rowShift;
        if (shiftedTop > XlsxSharp.XLHelper.MaxRowNumber)
        {
            return null;
        }

        int shiftedBottom = this.BottomRow + rowShift;
        if (shiftedBottom < XlsxSharp.XLHelper.MinRowNumber)
        {
            return null;
        }

        int clippedTop = Math.Max(shiftedTop, XlsxSharp.XLHelper.MinRowNumber);
        int clippedBottom = Math.Min(shiftedBottom, XlsxSharp.XLHelper.MaxRowNumber);

        return new Area(clippedTop, this.LeftColumn, clippedBottom, this.RightColumn);
    }

    /// <summary>
    /// Return a new range that has been shifted in horizontal direction by <paramref name="columnShift"/>.
    /// </summary>
    /// <param name="columnShift">By how much to shift the range, positive - rightward, negative - leftward.</param>
    /// <returns>Newly created area.</returns>
    internal Area ShiftColumns(int columnShift)
    {
        Point topLeftCorner = this.FirstPoint.ShiftColumn(columnShift);
        Point bottomRightCorner = this.LastPoint.ShiftColumn(columnShift);
        return new Area(topLeftCorner, bottomRightCorner);
    }

    /// <summary>
    /// Return a new range that has been shifted in horizontal direction by <paramref name="columnShift"/>.
    /// If the shifted area is out of sheet bounds, clip part that is out.
    /// </summary>
    /// <param name="columnShift">How many columns to shift.</param>
    /// <returns>Shifted clipped area or <c>null</c> if area was shifted completely out of a sheet.</returns>
    internal Area? ShiftColumnsAndClip(int columnShift)
    {
        int shiftedLeft = this.LeftColumn + columnShift;
        if (shiftedLeft > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            return null;
        }

        int shiftedRight = this.RightColumn + columnShift;
        if (shiftedRight < XlsxSharp.XLHelper.MinColumnNumber)
        {
            return null;
        }

        int clippedLeft = Math.Max(shiftedLeft, XlsxSharp.XLHelper.MinColumnNumber);
        int clippedRight = Math.Min(shiftedRight, XlsxSharp.XLHelper.MaxColumnNumber);

        return new Area(this.TopRow, clippedLeft, this.BottomRow, clippedRight);
    }

    public IEnumerator<Point> GetEnumerator()
    {
        for (int row = this.TopRow; row <= this.BottomRow; ++row)
        {
            for (int col = this.LeftColumn; col <= this.RightColumn; ++col)
            {
                yield return new Point(row, col);
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    /// <summary>
    /// Calculate size and position of the area when another area is inserted into a sheet.
    /// </summary>
    /// <param name="insertedArea">Inserted area.</param>
    /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
    /// <returns><c>true</c> if results wasn't partially shifted.</returns>
    internal bool TryInsertAreaAndShiftRight(Area insertedArea, out Area? result)
    {
        // Inserted fully upward, downward or to the right
        if (
            insertedArea.BottomRow < this.TopRow
            || insertedArea.TopRow > this.BottomRow
            || insertedArea.LeftColumn > this.RightColumn
        )
        {
            result = this;
            return true;
        }

        bool fullyOverlaps =
            insertedArea.TopRow <= this.TopRow && insertedArea.BottomRow >= this.BottomRow;
        if (!fullyOverlaps)
        {
            result = null;
            return false;
        }

        // Are is effectively inserted into a seam at the left column of the insertedArea
        if (insertedArea.LeftColumn <= this.LeftColumn)
        {
            // Area is completely pushed out
            if (this.LeftColumn + insertedArea.Width > XlsxSharp.XLHelper.MaxColumnNumber)
            {
                result = null;
                return true;
            }

            // Area is partially pushed out
            if (this.RightColumn + insertedArea.Width > XlsxSharp.XLHelper.MaxColumnNumber)
            {
                int pushedOutColsCount =
                    this.RightColumn + insertedArea.Width - XlsxSharp.XLHelper.MaxColumnNumber;
                int keepCols = this.Width - pushedOutColsCount;
                Area resized = this.SliceFromLeft(keepCols);
                result = resized.ShiftColumns(insertedArea.Width);
                return true;
            }

            // Not pushed out = only shift
            result = this.ShiftColumns(insertedArea.Width);
            return true;
        }

        result = this.ExtendRight(insertedArea.Width);
        return true;
    }

    /// <summary>
    /// Calculate size and position of the area when another area is inserted into a sheet.
    /// </summary>
    /// <param name="insertedArea">Inserted area.</param>
    /// <param name="result">The result, might be <c>null</c> as a valid result if area is pushed out.</param>
    /// <returns><c>true</c> if results wasn't partially shifted.</returns>
    internal bool TryInsertAreaAndShiftDown(Area insertedArea, out Area? result)
    {
        // Inserted fully to the left, to the right or below
        if (
            insertedArea.RightColumn < this.LeftColumn
            || insertedArea.LeftColumn > this.RightColumn
            || insertedArea.TopRow > this.BottomRow
        )
        {
            result = this;
            return true;
        }

        bool fullyOverlaps =
            insertedArea.LeftColumn <= this.LeftColumn
            && insertedArea.RightColumn >= this.RightColumn;
        if (!fullyOverlaps)
        {
            result = null;
            return false;
        }

        // Are is effectively inserted into a seam at the top row of the insertedArea
        if (insertedArea.TopRow <= this.TopRow)
        {
            // Area is completely pushed out
            if (this.TopRow + insertedArea.Height > XlsxSharp.XLHelper.MaxRowNumber)
            {
                result = null;
                return true;
            }

            // Area is partially pushed out
            if (this.BottomRow + insertedArea.Height > XlsxSharp.XLHelper.MaxRowNumber)
            {
                int pushedOutRowsCount =
                    this.BottomRow + insertedArea.Height - XlsxSharp.XLHelper.MaxRowNumber;
                int keepRows = this.Height - pushedOutRowsCount;
                Area resized = this.SliceFromTop(keepRows);
                result = resized.ShiftRows(insertedArea.Height);
                return true;
            }

            // Not pushed out = only shift
            result = this.ShiftRows(insertedArea.Height);
            return true;
        }

        result = this.ExtendBelow(insertedArea.Height);
        return true;
    }

    /// <summary>
    /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
    /// from sheet. If cells the left of the area are deleted, the area shifts to the left.
    /// If <paramref name="deletedArea"/> is within the area, the width of the area decreases.
    /// </summary>
    /// <remarks>
    /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
    /// decide what to do.
    /// </remarks>
    /// <returns>
    /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
    /// removed by <paramref name="deletedArea"/>.
    /// </returns>
    internal bool TryDeleteAreaAndShiftLeft(Area deletedArea, out Area? result)
    {
        // Deleted area is fully upwards, downwards or to the right of this area.
        if (
            deletedArea.BottomRow < this.TopRow
            || deletedArea.TopRow > this.BottomRow
            || deletedArea.LeftColumn > this.RightColumn
        )
        {
            result = this;
            return true;
        }

        bool coversWidth =
            deletedArea.LeftColumn <= this.LeftColumn
            && deletedArea.RightColumn >= this.RightColumn;
        bool coversHeight =
            deletedArea.TopRow <= this.TopRow && deletedArea.BottomRow >= this.BottomRow;
        bool fullyCovered = coversWidth && coversHeight;
        if (fullyCovered)
        {
            result = null;
            return true;
        }

        // When a slice form a top/bottom is deleted, the rest doesn't move.
        // There is no split either. Whole slice is just removed.
        bool deletedTopSlice =
            coversWidth
            && deletedArea.TopRow <= this.TopRow
            && deletedArea.BottomRow < this.BottomRow;
        if (deletedTopSlice)
        {
            int sliceRows = this.BottomRow - deletedArea.BottomRow;
            result = this.SliceFromBottom(sliceRows);
            return true;
        }

        bool deletedBottomSlice =
            coversWidth
            && deletedArea.BottomRow >= this.BottomRow
            && deletedArea.TopRow > this.TopRow;
        if (deletedBottomSlice)
        {
            int sliceRows = deletedArea.TopRow - this.TopRow;
            result = this.SliceFromTop(sliceRows);
            return true;
        }

        // Slice cases were already dealt with, anything that doesn't cover height would cause split
        if (!coversHeight)
        {
            result = null;
            return false;
        }

        bool deletesColumnsToLeft = deletedArea.LeftColumn < this.LeftColumn;
        bool deletesColumnsOfArea =
            deletedArea.LeftColumn <= this.RightColumn
            && deletedArea.RightColumn >= this.LeftColumn;
        Area repositioned = this;
        if (deletesColumnsOfArea)
        {
            // Decrease width of repositioned area
            int left = Math.Max(deletedArea.LeftColumn, repositioned.LeftColumn);
            int right = Math.Min(deletedArea.RightColumn, repositioned.RightColumn);

            int columnsToDelete = right - left + 1;
            int newWidth = repositioned.Width - columnsToDelete;
            if (newWidth == 0)
            {
                result = null;
                return true;
            }

            repositioned = repositioned.SliceFromLeft(newWidth);
        }

        if (deletesColumnsToLeft)
        {
            // There are some deleted columns to the left of the area -> shift left
            int deletedLastColumnsOutwards = Math.Min(
                repositioned.LeftColumn - 1,
                deletedArea.RightColumn
            );

            int shiftLeft = deletedLastColumnsOutwards - deletedArea.LeftColumn + 1;
            repositioned = repositioned.ShiftColumns(-shiftLeft);
        }

        result = repositioned;
        return true;
    }

    /// <summary>
    /// Take the area and reposition it as if the <paramref name="deletedArea"/> was removed
    /// from sheet. If cells upward of the area are deleted, the area shifts to the upward.
    /// If <paramref name="deletedArea"/> is within the area, the height of the area decreases.
    /// </summary>
    /// <remarks>
    /// If the method returns <c>false</c>, there is a partial cover and it's up to you to
    /// decide what to do.
    /// </remarks>
    /// <returns>
    /// The <paramref name="result"/> has a value <c>null</c> if the range was completely
    /// removed by <paramref name="deletedArea"/>.
    /// </returns>
    internal bool TryDeleteAreaAndShiftUp(Area deletedArea, out Area? result)
    {
        // Deleted area is fully on left, right or bottom side of this area.
        if (
            deletedArea.RightColumn < this.LeftColumn
            || deletedArea.LeftColumn > this.RightColumn
            || deletedArea.TopRow > this.BottomRow
        )
        {
            result = this;
            return true;
        }

        bool coversWidth =
            deletedArea.LeftColumn <= this.LeftColumn
            && deletedArea.RightColumn >= this.RightColumn;
        bool coversHeight =
            deletedArea.TopRow <= this.TopRow && deletedArea.BottomRow >= this.BottomRow;
        bool fullyCovered = coversWidth && coversHeight;
        if (fullyCovered)
        {
            result = null;
            return true;
        }

        // When a slice form a left/right is deleted, the rest doesn't move.
        // There is no split either. Whole slice is just removed.
        bool deletedLeftSlice =
            coversHeight
            && deletedArea.LeftColumn <= this.LeftColumn
            && deletedArea.RightColumn < this.RightColumn;
        if (deletedLeftSlice)
        {
            int sliceColumns = this.RightColumn - deletedArea.RightColumn;
            result = this.SliceFromRight(sliceColumns);
            return true;
        }

        bool deletedRightSlice =
            coversHeight
            && deletedArea.RightColumn >= this.RightColumn
            && deletedArea.LeftColumn > this.LeftColumn;
        if (deletedRightSlice)
        {
            int sliceRows = deletedArea.LeftColumn - this.LeftColumn;
            result = this.SliceFromLeft(sliceRows);
            return true;
        }

        // Slice cases were already dealt with, anything that doesn't cover height would cause split
        bool doesntOverlapWidth =
            deletedArea.LeftColumn > this.LeftColumn || deletedArea.RightColumn < this.RightColumn;
        if (doesntOverlapWidth)
        {
            result = null;
            return false;
        }

        bool deletesRowsAboveArea = deletedArea.TopRow < this.TopRow;
        bool deletesRowsOfArea =
            deletedArea.TopRow <= this.BottomRow && deletedArea.BottomRow >= this.TopRow;
        Area repositioned = this;
        if (deletesRowsOfArea)
        {
            // Decrease height of repositioned area
            int top = Math.Max(deletedArea.TopRow, repositioned.TopRow);
            int bottom = Math.Min(deletedArea.BottomRow, repositioned.BottomRow);

            int rowsToDelete = bottom - top + 1;
            int newHeight = repositioned.Height - rowsToDelete;
            if (newHeight == 0)
            {
                result = null;
                return true;
            }

            repositioned = repositioned.SliceFromTop(newHeight);
        }

        if (deletesRowsAboveArea)
        {
            // There are some deleted rows above the area -> shift up
            int deletedLastRowAboveArea = Math.Min(repositioned.TopRow - 1, deletedArea.BottomRow);

            int shiftUp = deletedLastRowAboveArea - deletedArea.TopRow + 1;
            repositioned = repositioned.ShiftRows(-shiftUp);
        }

        result = repositioned;
        return true;
    }

    /// <summary>
    /// Determine a areas that contain all cells of this area without <paramref name="range"/>
    /// and add them to the <paramref name="nonExcludedAreas"/>.
    /// </summary>
    /// <param name="range">Range to exclude from this one.</param>
    /// <param name="nonExcludedAreas">A list to which add remaining (non-excluded) areas.</param>
    /// <returns>If an area was excluded, the excluded area.</returns>
    internal Area? Exclude(Area range, List<Area> nonExcludedAreas)
    {
        if (this.Intersect(range) is not { } intersection)
        {
            nonExcludedAreas.Add(this);
            return null;
        }

        // top
        if (this.TopRow < intersection.TopRow)
        {
            nonExcludedAreas.Add(
                new Area(this.TopRow, this.LeftColumn, intersection.TopRow - 1, this.RightColumn)
            );
        }

        // bottom
        if (this.BottomRow > intersection.BottomRow)
        {
            nonExcludedAreas.Add(
                new Area(
                    intersection.BottomRow + 1,
                    this.LeftColumn,
                    this.BottomRow,
                    this.RightColumn
                )
            );
        }

        // left
        if (this.LeftColumn < intersection.LeftColumn)
        {
            nonExcludedAreas.Add(
                new Area(
                    intersection.TopRow,
                    this.LeftColumn,
                    intersection.BottomRow,
                    intersection.LeftColumn - 1
                )
            );
        }

        // right
        if (this.RightColumn > intersection.RightColumn)
        {
            nonExcludedAreas.Add(
                new Area(
                    intersection.TopRow,
                    intersection.RightColumn + 1,
                    intersection.BottomRow,
                    this.RightColumn
                )
            );
        }

        return intersection;
    }

    /// <summary>
    /// Return an area that has dimensions as if columns were inserted at <paramref name="insertedLeftColumn"/>.
    /// Mimics Excel behavior.
    /// </summary>
    /// <param name="insertedLeftColumn">A position where columns are inserted.</param>
    /// <param name="insertedWidth">How many columns were inserted.</param>
    internal Area? ShiftOrExtendRight(int insertedLeftColumn, int insertedWidth)
    {
        Debug.Assert(insertedWidth >= 0);

        // Area inserted at the right edge extends - that is the reason for - 1
        if (this.RightColumn < insertedLeftColumn - 1)
        {
            // inserted is to the right of area -> no shift
            return this;
        }

        if (this.LeftColumn >= insertedLeftColumn)
        {
            // Inserted is to the left of affected area -> shift
            return this.ShiftColumnsAndClip(insertedWidth);
        }

        // inserted is in the middle of affected: affectedLeft < insertedLeft <= affectedRight
        return this.ExtendRight(insertedWidth);
    }

    /// <summary>
    /// Return an area that has dimensions as if a rows were inserted at <paramref name="insertedTopRow"/>.
    /// Mimics Excel behavior.
    /// </summary>
    /// <param name="insertedTopRow">A position where rows are inserted.</param>
    /// <param name="insertedHeight">How many rows were inserted.</param>
    internal Area? ShiftOrExtendDown(int insertedTopRow, int insertedHeight)
    {
        Debug.Assert(insertedHeight >= 0);

        // Area inserted at the bottom edge extends - that is the reason for - 1
        if (this.BottomRow < insertedTopRow - 1)
        {
            // inserted is below the area -> no shift
            return this;
        }

        if (this.TopRow >= insertedTopRow)
        {
            // Inserted is above the area -> shift
            return this.ShiftRowsAndClip(insertedHeight);
        }

        // inserted is in the middle of affected: affectedTop < insertedTop <= affectedBottom
        return this.ExtendBelow(insertedHeight);
    }

    /// <summary>
    /// Return an area that has dimensions as if a rows were deleted from <paramref name="deletedTopRow"/>.
    /// Mimics Excel behavior.
    /// </summary>
    /// <param name="deletedTopRow">A position from which where rows are deleted.</param>
    /// <param name="deletedHeight">How many rows were deleted.</param>
    internal Area? ShiftOrShrinkUp(int deletedTopRow, int deletedHeight)
    {
        Debug.Assert(deletedHeight >= 0);
        if (this.BottomRow < deletedTopRow || deletedHeight == 0)
        {
            // deleted is below the area -> no shift or shrink
            return this;
        }

        int deletedBottomRow = deletedTopRow + deletedHeight - 1;
        if (deletedBottomRow < this.TopRow)
        {
            // Deleted area is completely above the area -> only shift
            return this.ShiftRows(-deletedHeight);
        }

        // Shrink by how much deletedArea and area overlap
        int shrink =
            Math.Min(this.BottomRow, deletedBottomRow) - Math.Max(this.TopRow, deletedTopRow) + 1;
        if (shrink == this.Height)
        {
            return null;
        }

        int shift = Math.Max(this.TopRow - deletedTopRow, 0);
        Area shifted = this.ShiftRows(-shift);
        return new Area(
            shifted.TopRow,
            shifted.LeftColumn,
            shifted.BottomRow - shrink,
            shifted.RightColumn
        );
    }

    /// <summary>
    /// Return an area that has dimensions as if a column were deleted from <paramref name="deletedLeftColumn"/>.
    /// Mimics Excel behavior.
    /// </summary>
    /// <param name="deletedLeftColumn">A position from which where columns are deleted.</param>
    /// <param name="deletedWidth">How many columns were deleted.</param>
    internal Area? ShiftOrShrinkLeft(int deletedLeftColumn, int deletedWidth)
    {
        Debug.Assert(deletedWidth >= 0);
        if (this.RightColumn < deletedLeftColumn || deletedWidth == 0)
        {
            // deleted is to the right of area -> no shift or shrink
            return this;
        }

        int deletedRightColumn = deletedLeftColumn + deletedWidth - 1;
        if (deletedRightColumn < this.LeftColumn)
        {
            // Deleted area is completely to left of area -> only shift
            return this.ShiftColumns(-deletedWidth);
        }

        // Shrink by how much deletedArea and area overlap
        int shrink =
            Math.Min(this.RightColumn, deletedRightColumn)
            - Math.Max(this.LeftColumn, deletedLeftColumn)
            + 1;
        if (shrink == this.Width)
        {
            return null;
        }

        int shift = Math.Max(this.LeftColumn - deletedLeftColumn, 0);
        Area shifted = this.ShiftColumns(-shift);
        return new Area(
            shifted.TopRow,
            shifted.LeftColumn,
            shifted.BottomRow,
            shifted.RightColumn - shrink
        );
    }

    /// <summary>
    /// Split the area above the <paramref name="row"/> and put result into the <paramref name="above"/> and
    /// <paramref name="below"/>.
    /// </summary>
    /// <returns><c>true</c> if <paramref name="above"/> is not null.</returns>
    internal bool SplitAbove(int row, [NotNullWhen(true)] out Area? above, out Area? below)
    {
        if (row is < XlsxSharp.XLHelper.MinRowNumber or > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        if (this.BottomRow < row)
        {
            above = this;
            below = null;
            return true;
        }

        if (this.TopRow >= row)
        {
            above = null;
            below = this;
            return false;
        }

        above = new Area(this.TopRow, this.LeftColumn, row - 1, this.RightColumn);
        below = new Area(row, this.LeftColumn, this.BottomRow, this.RightColumn);
        return true;
    }

    /// <summary>
    /// Split the area below the <paramref name="row"/> and put result into the <paramref name="below"/> and
    /// <paramref name="above"/>.
    /// </summary>
    /// <returns><c>true</c> if <paramref name="below"/> is not null.</returns>
    internal bool SplitBelow(int row, [NotNullWhen(true)] out Area? below, out Area? above)
    {
        if (row is < XlsxSharp.XLHelper.MinRowNumber or > XlsxSharp.XLHelper.MaxRowNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(row));
        }

        if (this.TopRow > row)
        {
            below = this;
            above = null;
            return true;
        }

        if (this.BottomRow <= row)
        {
            below = null;
            above = this;
            return false;
        }

        below = new Area(row + 1, this.LeftColumn, this.BottomRow, this.RightColumn);
        above = new Area(this.TopRow, this.LeftColumn, row, this.RightColumn);
        return true;
    }

    /// <summary>
    /// Split the area before the <paramref name="column"/> and put result into the <paramref name="left"/> and
    /// <paramref name="right"/>.
    /// </summary>
    /// <returns><c>true</c> if <paramref name="left"/> is not null.</returns>
    internal bool SplitBefore(int column, [NotNullWhen(true)] out Area? left, out Area? right)
    {
        if (column is < XlsxSharp.XLHelper.MinColumnNumber or > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(column));
        }

        if (this.RightColumn < column)
        {
            left = this;
            right = null;
            return true;
        }

        if (this.LeftColumn >= column)
        {
            left = null;
            right = this;
            return false;
        }

        left = new Area(this.TopRow, this.LeftColumn, this.BottomRow, column - 1);
        right = new Area(this.TopRow, column, this.BottomRow, this.RightColumn);
        return true;
    }

    /// <summary>
    /// Split the area after the <paramref name="column"/> and put result into the <paramref name="right"/> and
    /// <paramref name="left"/>.
    /// </summary>
    /// <returns><c>true</c> if <paramref name="right"/> is not null.</returns>
    internal bool SplitAfter(int column, [NotNullWhen(true)] out Area? right, out Area? left)
    {
        if (column is < XlsxSharp.XLHelper.MinColumnNumber or > XlsxSharp.XLHelper.MaxColumnNumber)
        {
            throw new ArgumentOutOfRangeException(nameof(column));
        }

        if (this.LeftColumn > column)
        {
            right = this;
            left = null;
            return true;
        }

        if (this.RightColumn <= column)
        {
            right = null;
            left = this;
            return false;
        }

        right = new Area(this.TopRow, column + 1, this.BottomRow, this.RightColumn);
        left = new Area(this.TopRow, this.LeftColumn, this.BottomRow, column);
        return true;
    }

    internal XLAreaList ToAreaList() => new(this);
}

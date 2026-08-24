using System.Diagnostics;

namespace XlsxSharp.Excel;

/// <summary>
/// An point (address) in a worksheet, an equivalent of <c>ST_CellRef</c>.
/// </summary>
/// <remarks>Unlike the XLAddress, sheet can never be invalid.</remarks>
[DebuggerDisplay("{XLHelper.GetColumnLetterFromNumber(Column)+Row}")]
internal readonly struct Point : IEquatable<Point>, IComparable<Point>
{
    public Point(int row, int column)
    {
        this.Row = row;
        this.Column = column;
    }

    /// <summary>
    /// 1-based row number in a sheet.
    /// </summary>
    public readonly int Row;

    /// <summary>
    /// 1-based column number in a sheet.
    /// </summary>
    public readonly int Column;

    public static implicit operator Area(Point point) => new(point);

    public override bool Equals(object? obj) => obj is Point point && this.Equals(point);

    public bool Equals(Point other) => this.Row == other.Row && this.Column == other.Column;

    public override int GetHashCode() => (this.Row * -1) ^ this.Column;

    public static bool operator ==(Point a, Point b) => a.Row == b.Row && a.Column == b.Column;

    public static bool operator !=(Point a, Point b) => a.Row != b.Row || a.Column != b.Column;

    /// <summary>
    /// Get offset that must be added to <paramref name="origin"/> so we can get <paramref name="target"/>.
    /// </summary>
    public static XLSheetOffset operator -(Point target, Point origin) =>
        new(target.Row - origin.Row, target.Column - origin.Column);

    /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
    public static Point Parse(string text) => Parse(text.AsSpan());

    /// <summary>
    /// Parse point per type <c>ST_CellRef</c> from
    /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db11a912-b1cb-4dff-b46d-9bedfd10cef0">2.1.1108 Part 4 Section 3.18.8, ST_CellRef (Cell Reference)</a>
    /// </summary>
    /// <param name="input">Input text</param>
    /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
    public static Point Parse(ReadOnlySpan<char> input)
    {
        if (!TryParse(input, out Point point))
        {
            throw new FormatException(
                $"Sheet point doesn't have correct format: '{input.ToString()}'."
            );
        }

        return point;
    }

    /// <summary>
    /// Try to parse sheet point. Doesn't accept any extra whitespace anywhere in the input.
    /// Letters must be upper case.
    /// </summary>
    public static bool TryParse(ReadOnlySpan<char> input, out Point point)
    {
        point = default;

        // Don't reuse inefficient logic from XLAddress
        if (input.Length < 2)
        {
            return false;
        }

        int i = 0;
        char c = input[i++];
        if (!IsLetter(c))
        {
            return false;
        }

        int columnIndex = c - 'A' + 1;
        while (i < input.Length && IsLetter(c = input[i]))
        {
            columnIndex = columnIndex * 26 + c - 'A' + 1;
            i++;
        }

        if (i > 3)
        {
            return false;
        }

        if (i == input.Length)
        {
            return false;
        }

        // Everything else must be digits
        c = input[i++];

        // First letter can't be 0
        if (c is < '1' or > '9')
        {
            return false;
        }

        int rowIndex = c - '0';
        while (i < input.Length && IsDigit(c = input[i]))
        {
            rowIndex = rowIndex * 10 + c - '0';
            i++;
        }

        if (i != input.Length)
        {
            return false;
        }

        if (
            rowIndex > XlsxSharp.XLHelper.MaxRowNumber
            || columnIndex > XlsxSharp.XLHelper.MaxColumnNumber
        )
        {
            return false;
        }

        point = new Point(rowIndex, columnIndex);
        return true;

        static bool IsLetter(char c) => c is >= 'A' and <= 'Z';
        static bool IsDigit(char c) => c is >= '0' and <= '9';
    }

    /// <summary>
    /// Write the sheet point as a reference to the span (e.g. <c>A1</c>).
    /// </summary>
    /// <param name="output">Must be at least 10 chars long</param>
    /// <returns>Number of chars </returns>
    public int Format(Span<char> output)
    {
        string columnLetters = XlsxSharp.XLHelper.GetColumnLetterFromNumber(this.Column);
        for (int i = 0; i < columnLetters.Length; ++i)
        {
            output[i] = columnLetters[i];
        }

        int digitCount = GetDigitCount(this.Row);
        int rowRemainder = this.Row;
        int formattedLength = digitCount + columnLetters.Length;
        for (int i = formattedLength - 1; i >= columnLetters.Length; --i)
        {
            int digit = rowRemainder % 10;
            rowRemainder /= 10;
            output[i] = (char)(digit + '0');
        }

        return formattedLength;
    }

    public override string ToString()
    {
        Span<char> text = stackalloc char[10];
        int len = this.Format(text);
        return text.Slice(0, len).ToString();
    }

    private static int GetDigitCount(int n)
    {
        if (n < 10L)
        {
            return 1;
        }

        if (n < 100L)
        {
            return 2;
        }

        if (n < 1000L)
        {
            return 3;
        }

        if (n < 10000L)
        {
            return 4;
        }

        if (n < 100000L)
        {
            return 5;
        }

        if (n < 1000000L)
        {
            return 6;
        }

        return 7; // Row can't have more digits
    }

    /// <summary>
    /// Create a sheet point from the address. Workbook is ignored.
    /// </summary>
    internal static Point FromAddress(IXLAddress address) =>
        new(address.RowNumber, address.ColumnNumber);

    internal static Point FromCell(IXLCell cell) => ((XLCell)cell).Point;

    public int CompareTo(Point other)
    {
        int rowComparison = this.Row.CompareTo(other.Row);
        if (rowComparison != 0)
        {
            return rowComparison;
        }

        return this.Column.CompareTo(other.Column);
    }

    /// <summary>
    /// Is the point within the range or below the range?
    /// </summary>
    internal bool InRangeOrBelow(in Area range) =>
        this.Row >= range.FirstPoint.Row
        && this.Column >= range.FirstPoint.Column
        && this.Column <= range.LastPoint.Column;

    /// <summary>
    /// Is the point within the range or to the right of the range?
    /// </summary>
    internal bool InRangeOrToRight(in Area range) =>
        this.Column >= range.FirstPoint.Column
        && this.Row >= range.FirstPoint.Row
        && this.Row <= range.LastPoint.Row;

    /// <summary>
    /// Return a new point that has its row coordinate shifted by <paramref name="rowShift"/>.
    /// </summary>
    /// <param name="rowShift">How many rows will new point be shifted. Positive - new point
    ///     is downwards, negative - new point is upwards relative to the current point.</param>
    /// <returns>Shifted point.</returns>
    internal Point ShiftRow(int rowShift) => new(this.Row + rowShift, this.Column);

    /// <summary>
    /// Return a new point that has its column coordinate shifted by <paramref name="columnShift"/>.
    /// </summary>
    /// <param name="columnShift">How many columns will new point be shifted. Positive - new
    ///     point is to the right, negative - new point is to the left.</param>
    /// <returns>Shifted point.</returns>
    internal Point ShiftColumn(int columnShift) => new(this.Row, this.Column + columnShift);
}

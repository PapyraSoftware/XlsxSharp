#nullable disable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XlsxSharp.Excel.CalcEngine;

/// <summary>
/// A base class for an 2D array. Every array is at least 1x1.
/// </summary>
internal abstract class Array : IEnumerable<ScalarValue>
{
    /// <summary>
    /// Width of the array, at least 1.
    /// </summary>
    public abstract int Width { get; }

    /// <summary>
    /// Height of the array, at least 1.
    /// </summary>
    public abstract int Height { get; }

    /// <summary>
    /// Get a value at specified coordinate.
    /// </summary>
    /// <param name="y">Uses 0-based notation.</param>
    /// <param name="x">Uses 0-based notation.</param>
    public abstract ScalarValue this[int y, int x] { get; }

    /// <summary>
    /// An iterator over all elements of an array, from top to bottom, from left to right.
    /// </summary>
    public virtual IEnumerator<ScalarValue> GetEnumerator() => this.FlattenArray().GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    protected IEnumerable<ScalarValue> FlattenArray()
    {
        for (int row = 0; row < this.Height; row++)
        {
            for (int col = 0; col < this.Width; col++)
            {
                yield return this[row, col];
            }
        }
    }

    /// <summary>
    /// Return a new array that was created by applying a function to each element of the array.
    /// </summary>
    public Array Apply(Func<ScalarValue, ScalarValue> op)
    {
        ScalarValue[,] data = new ScalarValue[this.Height, this.Width];
        for (int y = 0; y < this.Height; ++y)
        for (int x = 0; x < this.Width; ++x)
        {
            data[y, x] = op(this[y, x]);
        }

        return new ConstArray(data);
    }

    /// <summary>
    /// Return a new array that was created by applying a function to each element of the left and right array.
    /// Arrays can have different size and missing values are replaced by <c>#N/A</c>.
    /// </summary>
    public Array Apply(Array rightArray, BinaryFunc func, CalcContext ctx)
    {
        Array leftArray = this;
        int width = Math.Max(leftArray.Width, rightArray.Width);
        int height = Math.Max(leftArray.Height, rightArray.Height);
        ScalarValue[,] data = new ScalarValue[height, width];
        for (int y = 0; y < height; ++y)
        {
            for (int x = 0; x < width; ++x)
            {
                ScalarValue leftItem =
                    x < leftArray.Width && y < leftArray.Height
                        ? leftArray[y, x]
                        : XLError.NoValueAvailable;
                ScalarValue rightItem =
                    x < rightArray.Width && y < rightArray.Height
                        ? rightArray[y, x]
                        : XLError.NoValueAvailable;
                data[y, x] = func(leftItem, rightItem, ctx);
            }
        }

        return new ConstArray(data);
    }

    /// <summary>
    /// Broadcast array for calculation of array formulas.
    /// </summary>
    public Array Broadcast(int rows, int columns)
    {
        if (this.Width == columns && this.Height == rows)
        {
            return this;
        }

        if (this.Width == 1 && this.Height == 1)
        {
            return new ScalarArray(this[0, 0], columns, rows);
        }

        if (this.Width == 1)
        {
            return new RepeatedColumnArray(this, rows, columns);
        }

        if (this.Height == 1)
        {
            return new RepeatedRowArray(this, rows, columns);
        }

        return new ResizedArray(this, rows, columns);
    }
}

/// <summary>
/// An array of scalar values.
/// </summary>
internal class ConstArray : Array
{
    private readonly ScalarValue[,] _data;

    public ConstArray(ScalarValue[,] data)
    {
        if (data.GetLength(0) < 1 || data.GetLength(1) < 1)
        {
            throw new ArgumentException("Array must be at least 1x1.", nameof(data));
        }

        this._data = data;
    }

    public override ScalarValue this[int y, int x] => this._data[y, x];

    public override int Width => this._data.GetLength(1);

    public override int Height => this._data.GetLength(0);
}

/// <summary>
/// Array for array literal from a parser. It uses a 1D array of values as a storage.
/// </summary>
internal class LiteralArray : Array
{
    private readonly int _rows;
    private readonly int _columns;
    private readonly IReadOnlyList<ScalarValue> _elements;

    /// <summary>
    /// Create a new instance of a <see cref="LiteralArray"/>.
    /// </summary>
    /// <param name="rows">Number of rows of an array/</param>
    /// <param name="columns">Number of columns of an array.</param>
    /// <param name="elements">Row by row data of the array. Has the expected size of an array.</param>
    public LiteralArray(int rows, int columns, IReadOnlyList<ScalarValue> elements)
    {
        if (rows * columns != elements.Count)
        {
            throw new ArgumentException(
                "Number of elements in not the same as size of an array.",
                nameof(elements)
            );
        }

        this._rows = rows;
        this._columns = columns;
        this._elements = elements;
    }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (x < 0 || x >= this._columns)
            {
                throw new ArgumentOutOfRangeException(nameof(x));
            }

            return this._elements[y * this._columns + x];
        }
    }

    public override int Width => this._columns;

    public override int Height => this._rows;
}

/// <summary>
/// A special case of an array that is actually only numbers.
/// </summary>
internal class NumberArray : Array
{
    private readonly double[,] _data;

    public NumberArray(double[,] data) => this._data = data;

    public override ScalarValue this[int y, int x] => this._data[y, x];

    public override int Width => this._data.GetLength(1);

    public override int Height => this._data.GetLength(0);
}

/// <summary>
/// An array that retrieves its value directly from the worksheet without allocating extra memory.
/// </summary>
internal class ReferenceArray : Array
{
    private readonly XLRangeAddress _area;
    private readonly CalcContext _context;
    private readonly int _offsetColumn;
    private readonly int _offsetRow;

    public ReferenceArray(XLRangeAddress area, CalcContext context)
    {
        this._area = area;
        this._context = context;
        this._offsetColumn = this._area.FirstAddress.ColumnNumber;
        this._offsetRow = area.FirstAddress.RowNumber;
    }

    public override ScalarValue this[int y, int x] =>
        this._context.GetCellValue(
            this._area.Worksheet,
            y + this._offsetRow,
            x + this._offsetColumn
        );

    public override int Width => this._area.ColumnSpan;

    public override int Height => this._area.RowSpan;
}

internal class RepeatedColumnArray : Array
{
    private readonly Array _columnArray;

    public RepeatedColumnArray(Array oneColumnArray, int rows, int columns)
    {
        Debug.Assert(oneColumnArray.Width == 1);
        this._columnArray = oneColumnArray;
        this.Width = columns;
        this.Height = rows;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int row, int column]
    {
        get
        {
            if (row >= this.Height || column >= this.Width)
            {
                throw new IndexOutOfRangeException();
            }

            if (row >= this._columnArray.Height)
            {
                return XLError.NoValueAvailable;
            }

            return this._columnArray[row, 0];
        }
    }
}

internal class RepeatedRowArray : Array
{
    private readonly Array _rowArray;

    internal RepeatedRowArray(Array oneRowArray, int rows, int columns)
    {
        Debug.Assert(oneRowArray.Height == 1);
        this._rowArray = oneRowArray;
        this.Width = columns;
        this.Height = rows;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int row, int column]
    {
        get
        {
            if (row >= this.Height || column >= this.Width)
            {
                throw new IndexOutOfRangeException();
            }

            if (column >= this._rowArray.Width)
            {
                return XLError.NoValueAvailable;
            }

            return this._rowArray[0, column];
        }
    }
}

/// <summary>
/// A resize array from another array. Extra items without value have <c>#N/A</c>.
/// </summary>
internal class ResizedArray : Array
{
    private readonly Array _original;

    public ResizedArray(Array original, int rows, int columns)
    {
        this._original = original;
        this.Height = rows;
        this.Width = columns;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (y >= this.Height || x >= this.Width)
            {
                throw new IndexOutOfRangeException();
            }

            return y < this._original.Height && x < this._original.Width
                ? this._original[y, x]
                : XLError.NoValueAvailable;
        }
    }
}

/// <summary>
/// An array where all elements have same value.
/// </summary>
internal class ScalarArray : Array
{
    private readonly ScalarValue _value;
    private readonly int _columns;
    private readonly int _rows;

    public ScalarArray(ScalarValue value, int columns, int rows)
    {
        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columns));
        }

        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rows));
        }

        this._value = value;
        this._columns = columns;
        this._rows = rows;
    }

    public override int Width => this._columns;

    public override int Height => this._rows;

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (x < 0 || x >= this._columns || y < 0 || y >= this._rows)
            {
                throw new IndexOutOfRangeException();
            }

            return this._value;
        }
    }

    public override IEnumerator<ScalarValue> GetEnumerator() =>
        Enumerable.Range(0, this._columns * this._rows).Select(_ => this._value).GetEnumerator();
}

internal class TransposedArray : Array
{
    private readonly Array _original;

    public TransposedArray(Array original) => this._original = original;

    public override ScalarValue this[int y, int x] => this._original[x, y];

    public override int Width => this._original.Height;

    public override int Height => this._original.Width;
}

/// <summary>
/// An array that is a rectangular slice of the original array.
/// </summary>
internal class SlicedArray : Array
{
    private readonly Array _original;
    private readonly int _rowOfs;
    private readonly int _colOfs;

    /// <summary>
    /// Create a sliced array from the original array.
    /// </summary>
    /// <param name="original">Original array.</param>
    /// <param name="rowOfs">The row offset indicating the starting row of the slice in the original array.</param>
    /// <param name="rows">The number of rows in the sliced array.</param>
    /// <param name="colOfs">The column offset indicating the starting column of the slice in the original array.</param>
    /// <param name="cols">The number of columns in the sliced array.</param>
    public SlicedArray(Array original, int rowOfs, int rows, int colOfs, int cols)
    {
        if (
            rowOfs < 0
            || rows < 1
            || colOfs < 0
            || cols < 1
            || rowOfs + rows > original.Height
            || colOfs + cols > original.Width
        )
        {
            throw new ArgumentOutOfRangeException();
        }

        this._original = original;
        this._rowOfs = rowOfs;
        this.Height = rows;
        this._colOfs = colOfs;
        this.Width = cols;
    }

    public override ScalarValue this[int y, int x] =>
        this._original[y + this._rowOfs, x + this._colOfs];

    public override int Width { get; }

    public override int Height { get; }
}

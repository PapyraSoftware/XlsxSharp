#nullable disable

using System.Diagnostics;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal struct XLAddress : IXLAddress, IEquatable<XLAddress>
{
    #region Static
    /// <summary>
    /// Create address without worksheet. For calculation only!
    /// </summary>
    /// <param name="cellAddressString"></param>
    public static XLAddress Create(string cellAddressString) => Create(null, cellAddressString);

    public static XLAddress Create(XLWorksheet worksheet, string cellAddressString)
    {
        bool fixedColumn = cellAddressString[0] == '$';
        int startPos;
        if (fixedColumn)
        {
            startPos = 1;
        }
        else
        {
            startPos = 0;
        }

        int rowPos = startPos;
        while (cellAddressString[rowPos] > '9')
        {
            rowPos++;
        }

        bool fixedRow = cellAddressString[rowPos] == '$';
        string columnLetter;
        int rowNumber;
        if (fixedRow)
        {
            if (fixedColumn)
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
            }
            else
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos);
            }

            rowNumber = int.Parse(
                cellAddressString.AsSpan(rowPos + 1),
                XlsxSharp.XLHelper.NumberStyle,
                XlsxSharp.XLHelper.ParseCulture
            );
        }
        else
        {
            if (fixedColumn)
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
            }
            else
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos);
            }

            rowNumber = int.Parse(
                cellAddressString.AsSpan(rowPos),
                XlsxSharp.XLHelper.NumberStyle,
                XlsxSharp.XLHelper.ParseCulture
            );
        }
        return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
    }

    #endregion Static

    #region Private fields

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private bool _fixedRow;

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private bool _fixedColumn;

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly int _rowNumber;

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly int _columnNumber;

    private string _trimmedAddress;

    #endregion Private fields

    #region Constructors

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
    /// </summary>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnLetter">The column letter of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    public XLAddress(int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
        : this(null, rowNumber, columnLetter, fixedRow, fixedColumn) { }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
    /// </summary>
    /// <param name = "worksheet"></param>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnLetter">The column letter of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    public XLAddress(
        XLWorksheet worksheet,
        int rowNumber,
        string columnLetter,
        bool fixedRow,
        bool fixedColumn
    )
        : this(
            worksheet,
            rowNumber,
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(columnLetter),
            fixedRow,
            fixedColumn
        ) { }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation. Attention: without worksheet for calculation only!
    /// </summary>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnNumber">The column number of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    public XLAddress(int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
        : this(null, rowNumber, columnNumber, fixedRow, fixedColumn) { }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation.
    /// </summary>
    /// <param name = "worksheet"></param>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnNumber">The column number of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    public XLAddress(
        XLWorksheet worksheet,
        int rowNumber,
        int columnNumber,
        bool fixedRow,
        bool fixedColumn
    )
        : this()
    {
        this.Worksheet = worksheet;

        this._rowNumber = rowNumber;
        this._columnNumber = columnNumber;
        this._fixedColumn = fixedColumn;
        this._fixedRow = fixedRow;
    }

    #endregion Constructors

    #region Properties

    public XLWorksheet Worksheet { get; internal set; }

    IXLWorksheet IXLAddress.Worksheet
    {
        [DebuggerStepThrough]
        get => this.Worksheet;
    }

    public bool HasWorksheet
    {
        [DebuggerStepThrough]
        get => this.Worksheet != null;
    }

    public bool FixedRow => this._fixedRow;

    public bool FixedColumn => this._fixedColumn;

    /// <summary>
    /// Gets the row number of this address.
    /// </summary>
    public int RowNumber => this._rowNumber;

    /// <summary>
    /// Gets the column number of this address.
    /// </summary>
    public int ColumnNumber => this._columnNumber;

    /// <summary>
    /// Gets the column letter(s) of this address.
    /// </summary>
    public string ColumnLetter => XlsxSharp.XLHelper.GetColumnLetterFromNumber(this._columnNumber);

    #endregion Properties

    #region Overrides

    public override string ToString()
    {
        if (!this.IsValid)
        {
            return "#REF!";
        }

        string retVal = this.ColumnLetter;
        if (this._fixedColumn)
        {
            retVal = "$" + retVal;
        }
        if (this._fixedRow)
        {
            retVal += "$";
        }
        retVal += this._rowNumber.ToInvariantString();
        return retVal;
    }

    public string ToString(XLReferenceStyle referenceStyle) => this.ToString(referenceStyle, false);

    public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;
        if (!this.IsValid)
        {
            address = "#REF!";
        }
        else if (referenceStyle == XLReferenceStyle.A1)
        {
            address = this.GetTrimmedAddress();
        }
        else if (
            referenceStyle == XLReferenceStyle.R1C1
            || this.HasWorksheet && this.Worksheet.Workbook.ReferenceStyle == XLReferenceStyle.R1C1
        )
        {
            address =
                "R"
                + this._rowNumber.ToInvariantString()
                + "C"
                + this.ColumnNumber.ToInvariantString();
        }
        else
        {
            address = this.GetTrimmedAddress();
        }

        if (includeSheet)
        {
            return string.Concat(
                this.WorksheetIsDeleted ? "#REF" : this.Worksheet.Name.EscapeSheetName(),
                '!',
                address
            );
        }

        return address;
    }

    #endregion Overrides

    #region Methods

    public string GetTrimmedAddress() =>
        this._trimmedAddress
        ?? (this._trimmedAddress = this.ColumnLetter + this._rowNumber.ToInvariantString());

    #endregion Methods

    #region Operator Overloads

    public static XLAddress operator +(XLAddress left, XLAddress right) =>
        new(
            left.Worksheet,
            left.RowNumber + right.RowNumber,
            left.ColumnNumber + right.ColumnNumber,
            left._fixedRow,
            left._fixedColumn
        );

    public static XLAddress operator -(XLAddress left, XLAddress right) =>
        new(
            left.Worksheet,
            left.RowNumber - right.RowNumber,
            left.ColumnNumber - right.ColumnNumber,
            left._fixedRow,
            left._fixedColumn
        );

    public static XLAddress operator +(XLAddress left, int right) =>
        new(
            left.Worksheet,
            left.RowNumber + right,
            left.ColumnNumber + right,
            left._fixedRow,
            left._fixedColumn
        );

    public static XLAddress operator -(XLAddress left, int right) =>
        new(
            left.Worksheet,
            left.RowNumber - right,
            left.ColumnNumber - right,
            left._fixedRow,
            left._fixedColumn
        );

    public static bool operator ==(XLAddress left, XLAddress right) => left.Equals(right);

    public static bool operator !=(XLAddress left, XLAddress right) => !(left == right);

    #endregion Operator Overloads

    #region Interface Requirements

    #region IEqualityComparer<XLCellAddress> Members

    public bool Equals(IXLAddress x, IXLAddress y) => x == y;

    public static new bool Equals(object x, object y) => x == y;

    #endregion IEqualityComparer<XLCellAddress> Members

    #region IEquatable<XLCellAddress> Members

    public bool Equals(IXLAddress other)
    {
        if (other == null)
        {
            return false;
        }

        return this._rowNumber == other.RowNumber
            && this._columnNumber == other.ColumnNumber
            && this._fixedRow == other.FixedRow
            && this._fixedColumn == other.FixedColumn;
    }

    public bool Equals(XLAddress other) =>
        this._rowNumber == other._rowNumber
        && this._columnNumber == other._columnNumber
        && this._fixedRow == other._fixedRow
        && this._fixedColumn == other._fixedColumn;

    public override bool Equals(object other) => this.Equals(other as IXLAddress);

    public override int GetHashCode()
    {
        int hashCode = 2122234362;
        hashCode = hashCode * -1521134295 + this._fixedRow.GetHashCode();
        hashCode = hashCode * -1521134295 + this._fixedColumn.GetHashCode();
        hashCode = hashCode * -1521134295 + this._rowNumber.GetHashCode();
        hashCode = hashCode * -1521134295 + this._columnNumber.GetHashCode();
        return hashCode;
    }

    public int GetHashCode(IXLAddress obj) => ((XLAddress)obj).GetHashCode();

    #endregion IEquatable<XLCellAddress> Members

    #endregion Interface Requirements

    public string ToStringRelative() => this.ToStringRelative(false);

    public string ToStringFixed() => this.ToStringFixed(XLReferenceStyle.Default);

    public string ToStringRelative(bool includeSheet)
    {
        string address = this.IsValid ? this.GetTrimmedAddress() : "#REF!";

        if (includeSheet)
        {
            return string.Concat(
                this.WorksheetIsDeleted ? "#REF" : this.Worksheet.Name.EscapeSheetName(),
                '!',
                address
            );
        }

        return address;
    }

    internal XLAddress WithoutWorksheet() =>
        new(this.RowNumber, this.ColumnNumber, this.FixedRow, this.FixedColumn);

    internal XLAddress WithWorksheet(XLWorksheet worksheet) =>
        new(worksheet, this.RowNumber, this.ColumnNumber, this.FixedRow, this.FixedColumn);

    public string ToStringFixed(XLReferenceStyle referenceStyle) =>
        this.ToStringFixed(referenceStyle, false);

    public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;

        if (referenceStyle == XLReferenceStyle.Default && this.HasWorksheet)
        {
            referenceStyle = this.Worksheet.Workbook.ReferenceStyle;
        }

        if (referenceStyle == XLReferenceStyle.Default)
        {
            referenceStyle = XLReferenceStyle.A1;
        }

        Debug.Assert(referenceStyle != XLReferenceStyle.Default);

        if (!this.IsValid)
        {
            address = "#REF!";
        }
        else
        {
            switch (referenceStyle)
            {
                case XLReferenceStyle.A1:
                    address = string.Concat(
                        '$',
                        this.ColumnLetter,
                        '$',
                        this._rowNumber.ToInvariantString()
                    );
                    break;

                case XLReferenceStyle.R1C1:
                    address = string.Concat(
                        'R',
                        this._rowNumber.ToInvariantString(),
                        'C',
                        this.ColumnNumber
                    );
                    break;

                default:
                    throw new NotImplementedException();
            }
        }

        if (includeSheet)
        {
            return string.Concat(
                this.WorksheetIsDeleted ? "#REF" : this.Worksheet.Name.EscapeSheetName(),
                '!',
                address
            );
        }

        return address;
    }

    public string UniqueId =>
        this.RowNumber.ToString("0000000") + this.ColumnNumber.ToString("00000");

    public bool IsValid =>
        0 < this.RowNumber
        && this.RowNumber <= XlsxSharp.XLHelper.MaxRowNumber
        && 0 < this.ColumnNumber
        && this.ColumnNumber <= XlsxSharp.XLHelper.MaxColumnNumber;

    private bool WorksheetIsDeleted => this.Worksheet?.IsDeleted == true;
}

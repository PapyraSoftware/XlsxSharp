using System;
using System.Collections.Generic;
using System.Diagnostics;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

internal readonly struct XLRangeAddress : IXLRangeAddress, IEquatable<XLRangeAddress>
{
    #region Static members

    public static XLRangeAddress EntireColumn(XLWorksheet worksheet, int column)
    {
        return new XLRangeAddress(
            new XLAddress(worksheet, 1, column, false, false),
            new XLAddress(worksheet, XlsxSharp.XLHelper.MaxRowNumber, column, false, false)
        );
    }

    public static XLRangeAddress EntireRow(XLWorksheet worksheet, int row)
    {
        return new XLRangeAddress(
            new XLAddress(worksheet, row, 1, false, false),
            new XLAddress(worksheet, row, XlsxSharp.XLHelper.MaxColumnNumber, false, false)
        );
    }

    public static readonly XLRangeAddress Invalid = new(
        new XLAddress(-1, -1, fixedRow: true, fixedColumn: true),
        new XLAddress(-1, -1, fixedRow: true, fixedColumn: true)
    );

    internal static XLRangeAddress FromSheetRange(XLWorksheet? worksheet, Area range)
    {
        return new XLRangeAddress(
            new XLAddress(
                worksheet,
                range.FirstPoint.Row,
                range.FirstPoint.Column,
                fixedRow: false,
                fixedColumn: false
            ),
            new XLAddress(
                range.LastPoint.Row,
                range.LastPoint.Column,
                fixedRow: false,
                fixedColumn: false
            )
        );
    }

    #endregion Static members

    #region Constructor

    public XLRangeAddress(XLAddress firstAddress, XLAddress lastAddress)
        : this()
    {
        this.Worksheet = firstAddress.Worksheet;
        this.FirstAddress = firstAddress;
        this.LastAddress = lastAddress;
    }

    public XLRangeAddress(XLWorksheet? worksheet, String rangeAddress)
        : this()
    {
        string addressToUse = rangeAddress.Contains('!')
            ? rangeAddress.Substring(rangeAddress.LastIndexOf('!') + 1)
            : rangeAddress;

        string firstPart;
        string secondPart;
        if (addressToUse.Contains(':'))
        {
            string[] arrRange = addressToUse.Split(':');
            firstPart = arrRange[0];
            secondPart = arrRange[1];
        }
        else
        {
            firstPart = addressToUse;
            secondPart = addressToUse;
        }

        if (XlsxSharp.XLHelper.IsValidA1Address(firstPart))
        {
            this.FirstAddress = XLAddress.Create(worksheet, firstPart);
            this.LastAddress = XLAddress.Create(worksheet, secondPart);
        }
        else
        {
            firstPart = firstPart.Replace("$", String.Empty);
            secondPart = secondPart.Replace("$", String.Empty);
            if (char.IsDigit(firstPart[0]))
            {
                this.FirstAddress = XLAddress.Create(worksheet, "A" + firstPart);
                this.LastAddress = XLAddress.Create(
                    worksheet,
                    XlsxSharp.XLHelper.MaxColumnLetter + secondPart
                );
            }
            else
            {
                this.FirstAddress = XLAddress.Create(worksheet, firstPart + "1");
                this.LastAddress = XLAddress.Create(
                    worksheet,
                    secondPart + XlsxSharp.XLHelper.MaxRowNumber.ToInvariantString()
                );
            }
        }

        this.Worksheet = worksheet;
    }

    #endregion Constructor

    #region Public properties

    public XLWorksheet? Worksheet { get; }

    public XLAddress FirstAddress { get; }

    public XLAddress LastAddress { get; }

    IXLWorksheet? IXLRangeAddress.Worksheet
    {
        get { return this.Worksheet; }
    }

    IXLAddress IXLRangeAddress.FirstAddress
    {
        [DebuggerStepThrough]
        get { return this.FirstAddress; }
    }

    IXLAddress IXLRangeAddress.LastAddress
    {
        [DebuggerStepThrough]
        get { return this.LastAddress; }
    }

    public bool IsValid => this.FirstAddress.IsValid && this.LastAddress.IsValid;

    public int ColumnSpan
    {
        get
        {
            if (!this.IsValid)
            {
                throw new InvalidOperationException("Range address is invalid.");
            }

            return Math.Abs(this.LastAddress.ColumnNumber - this.FirstAddress.ColumnNumber) + 1;
        }
    }

    public int NumberOfCells => this.ColumnSpan * this.RowSpan;

    internal long Size => this.ColumnSpan * (long)this.RowSpan;

    public int RowSpan
    {
        get
        {
            if (!this.IsValid)
            {
                throw new InvalidOperationException("Range address is invalid.");
            }

            return Math.Abs(this.LastAddress.RowNumber - this.FirstAddress.RowNumber) + 1;
        }
    }

    private bool WorksheetIsDeleted => this.Worksheet?.IsDeleted == true;

    #endregion Public properties

    #region Public methods

    public Boolean IsNormalized =>
        this.LastAddress.RowNumber >= this.FirstAddress.RowNumber
        && this.LastAddress.ColumnNumber >= this.FirstAddress.ColumnNumber;

    /// <summary>
    /// Lead a range address to a normal form - when <see cref="FirstAddress"/> points to the top-left address and
    /// <see cref="LastAddress"/> points to the bottom-right address.
    /// </summary>
    public XLRangeAddress Normalize()
    {
        if (
            this.FirstAddress.RowNumber <= this.LastAddress.RowNumber
            && this.FirstAddress.ColumnNumber <= this.LastAddress.ColumnNumber
        )
        {
            return this;
        }

        int firstRow,
            firstColumn,
            lastRow,
            lastColumn;
        bool firstRowFixed,
            firstColumnFixed,
            lastRowFixed,
            lastColumnFixed;

        if (this.FirstAddress.RowNumber <= this.LastAddress.RowNumber)
        {
            firstRow = this.FirstAddress.RowNumber;
            firstRowFixed = this.FirstAddress.FixedRow;
            lastRow = this.LastAddress.RowNumber;
            lastRowFixed = this.LastAddress.FixedRow;
        }
        else
        {
            firstRow = this.LastAddress.RowNumber;
            firstRowFixed = this.LastAddress.FixedRow;
            lastRow = this.FirstAddress.RowNumber;
            lastRowFixed = this.FirstAddress.FixedRow;
        }

        if (this.FirstAddress.ColumnNumber <= this.LastAddress.ColumnNumber)
        {
            firstColumn = this.FirstAddress.ColumnNumber;
            firstColumnFixed = this.FirstAddress.FixedColumn;
            lastColumn = this.LastAddress.ColumnNumber;
            lastColumnFixed = this.LastAddress.FixedColumn;
        }
        else
        {
            firstColumn = this.LastAddress.ColumnNumber;
            firstColumnFixed = this.LastAddress.FixedColumn;
            lastColumn = this.FirstAddress.ColumnNumber;
            lastColumnFixed = this.FirstAddress.FixedColumn;
        }

        return new XLRangeAddress(
            new XLAddress(
                this.FirstAddress.Worksheet,
                firstRow,
                firstColumn,
                firstRowFixed,
                firstColumnFixed
            ),
            new XLAddress(
                this.LastAddress.Worksheet,
                lastRow,
                lastColumn,
                lastRowFixed,
                lastColumnFixed
            )
        );
    }

    public bool Intersects(IXLRangeAddress otherAddress)
    {
        XLRangeAddress xlOtherAddress = (XLRangeAddress)otherAddress;
        return this.Intersects(in xlOtherAddress);
    }

    internal bool Intersects(in XLRangeAddress otherAddress)
    {
        return !( // See if the two ranges intersect...
            otherAddress.FirstAddress.ColumnNumber > this.LastAddress.ColumnNumber
            || otherAddress.LastAddress.ColumnNumber < this.FirstAddress.ColumnNumber
            || otherAddress.FirstAddress.RowNumber > this.LastAddress.RowNumber
            || otherAddress.LastAddress.RowNumber < this.FirstAddress.RowNumber
        );
    }

    public bool Contains(IXLAddress address)
    {
        XLAddress xlAddress = (XLAddress)address;
        return this.Contains(in xlAddress);
    }

    /// <summary>
    /// Does this range contains whole another range?
    /// </summary>
    public bool ContainsWhole(IXLRangeAddress range)
    {
        if (!range.IsValid)
        {
            return false;
        }

        return range.FirstAddress.ColumnNumber >= this.FirstAddress.ColumnNumber
            && range.FirstAddress.RowNumber >= this.FirstAddress.RowNumber
            && range.LastAddress.ColumnNumber <= this.LastAddress.ColumnNumber
            && range.LastAddress.RowNumber <= this.LastAddress.RowNumber;
    }

    internal IXLRangeAddress WithoutWorksheet()
    {
        return new XLRangeAddress(
            this.FirstAddress.WithoutWorksheet(),
            this.LastAddress.WithoutWorksheet()
        );
    }

    internal XLRangeAddress WithWorksheet(XLWorksheet worksheet)
    {
        return new XLRangeAddress(
            this.FirstAddress.WithWorksheet(worksheet),
            this.LastAddress.WithWorksheet(worksheet)
        );
    }

    internal bool Contains(in XLAddress address)
    {
        return this.FirstAddress.RowNumber <= address.RowNumber
            && address.RowNumber <= this.LastAddress.RowNumber
            && this.FirstAddress.ColumnNumber <= address.ColumnNumber
            && address.ColumnNumber <= this.LastAddress.ColumnNumber;
    }

    public String ToStringRelative()
    {
        return this.ToStringRelative(false);
    }

    public String ToStringFixed()
    {
        return this.ToStringFixed(XLReferenceStyle.A1);
    }

    public String ToStringRelative(Boolean includeSheet)
    {
        string address;
        if (!this.IsValid)
        {
            address = "#REF!";
        }
        else
        {
            if (this.IsEntireSheet())
            {
                address = $"1:{XlsxSharp.XLHelper.MaxRowNumber}";
            }
            else if (this.IsEntireRow())
            {
                address = String.Concat(
                    this.FirstAddress.RowNumber.ToString(),
                    ":",
                    this.LastAddress.RowNumber.ToString()
                );
            }
            else if (this.IsEntireColumn())
            {
                address = String.Concat(
                    this.FirstAddress.ColumnLetter,
                    ":",
                    this.LastAddress.ColumnLetter
                );
            }
            else
            {
                address = String.Concat(
                    this.FirstAddress.ToStringRelative(),
                    ":",
                    this.LastAddress.ToStringRelative()
                );
            }
        }

        if (includeSheet || this.WorksheetIsDeleted)
        {
            return String.Concat(
                this.WorksheetIsDeleted ? "#REF" : this.Worksheet!.Name.EscapeSheetName(),
                "!",
                address
            );
        }

        return address;
    }

    public String ToStringFixed(XLReferenceStyle referenceStyle)
    {
        return this.ToStringFixed(referenceStyle, false);
    }

    public String ToStringFixed(XLReferenceStyle referenceStyle, Boolean includeSheet)
    {
        string address;
        if (!this.IsValid)
        {
            address = "#REF!";
        }
        else
        {
            if (this.IsEntireSheet())
            {
                address = $"$1:${XlsxSharp.XLHelper.MaxRowNumber}";
            }
            else if (this.IsEntireRow())
            {
                address = String.Concat(
                    "$",
                    this.FirstAddress.RowNumber.ToString(),
                    ":$",
                    this.LastAddress.RowNumber.ToString()
                );
            }
            else if (this.IsEntireColumn())
            {
                address = String.Concat(
                    "$",
                    this.FirstAddress.ColumnLetter,
                    ":$",
                    this.LastAddress.ColumnLetter
                );
            }
            else
            {
                address = String.Concat(
                    this.FirstAddress.ToStringFixed(referenceStyle),
                    ":",
                    this.LastAddress.ToStringFixed(referenceStyle)
                );
            }
        }

        if (includeSheet || this.WorksheetIsDeleted)
        {
            return String.Concat(
                this.WorksheetIsDeleted ? "#REF" : this.Worksheet!.Name.EscapeSheetName(),
                "!",
                address
            );
        }

        return address;
    }

    public override string ToString()
    {
        if (!this.IsValid || this.WorksheetIsDeleted)
        {
            string worksheet = this.WorksheetIsDeleted ? "#REF!" : "";

            string address =
                (!this.FirstAddress.IsValid || !this.LastAddress.IsValid)
                    ? "#REF!"
                    : String.Concat(this.FirstAddress.ToString(), ":", this.LastAddress.ToString());
            return String.Concat(worksheet, address);
        }

        if (this.IsEntireSheet())
        {
            string worksheet = this.WorksheetIsDeleted ? "#REF!" : "";
            string address = $"1:{XlsxSharp.XLHelper.MaxRowNumber}";
            return String.Concat(worksheet, address);
        }
        else if (this.IsEntireRow())
        {
            string worksheet = this.WorksheetIsDeleted ? "#REF!" : "";
            string firstAddress = this.FirstAddress.IsValid
                ? this.FirstAddress.RowNumber.ToString()
                : "#REF!";
            string lastAddress = this.LastAddress.IsValid
                ? this.LastAddress.RowNumber.ToString()
                : "#REF!";

            return String.Concat(worksheet, firstAddress, ':', lastAddress);
        }
        else if (this.IsEntireColumn())
        {
            string worksheet = this.WorksheetIsDeleted ? "#REF!" : "";
            string firstAddress = this.FirstAddress.IsValid
                ? this.FirstAddress.ColumnLetter
                : "#REF!";
            string lastAddress = this.LastAddress.IsValid ? this.LastAddress.ColumnLetter : "#REF!";

            return String.Concat(worksheet, firstAddress, ':', lastAddress);
        }
        else
        {
            return String.Concat(this.FirstAddress.ToString(), ":", this.LastAddress.ToString());
        }
    }

    public string ToString(XLReferenceStyle referenceStyle)
    {
        return this.ToString(referenceStyle, false);
    }

    public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        if (referenceStyle == XLReferenceStyle.R1C1)
        {
            return this.ToStringFixed(referenceStyle, true);
        }
        else
        {
            return this.ToStringRelative(includeSheet);
        }
    }

    public override bool Equals(object obj)
    {
        if (!(obj is XLRangeAddress))
        {
            return false;
        }

        XLRangeAddress address = (XLRangeAddress)obj;
        return this.FirstAddress.Equals(address.FirstAddress)
            && this.LastAddress.Equals(address.LastAddress)
            && EqualityComparer<XLWorksheet?>.Default.Equals(this.Worksheet, address.Worksheet);
    }

    public override int GetHashCode()
    {
        int hashCode = -778064135;
        hashCode = hashCode * -1521134295 + this.FirstAddress.GetHashCode();
        hashCode = hashCode * -1521134295 + this.LastAddress.GetHashCode();
        hashCode =
            hashCode * -1521134295
            + EqualityComparer<XLWorksheet?>.Default.GetHashCode(this.Worksheet);
        return hashCode;
    }

    public bool Equals(XLRangeAddress other)
    {
        return ReferenceEquals(this.Worksheet, other.Worksheet)
            && this.FirstAddress == other.FirstAddress
            && this.LastAddress == other.LastAddress;
    }

    public bool IsSingleCell()
    {
        return this.IsValid
            && this.FirstAddress.RowNumber == this.LastAddress.RowNumber
            && this.FirstAddress.ColumnNumber == this.LastAddress.ColumnNumber;
    }

    public bool IsEntireColumn()
    {
        return this.IsValid
            && this.FirstAddress.RowNumber == 1
            && this.LastAddress.RowNumber == XlsxSharp.XLHelper.MaxRowNumber;
    }

    public bool IsEntireRow()
    {
        return this.IsValid
            && this.FirstAddress.ColumnNumber == 1
            && this.LastAddress.ColumnNumber == XlsxSharp.XLHelper.MaxColumnNumber;
    }

    public bool IsEntireSheet()
    {
        return this.IsValid && this.IsEntireColumn() && this.IsEntireRow();
    }

    public IXLRangeAddress Relative(
        IXLRangeAddress sourceRangeAddress,
        IXLRangeAddress targetRangeAddress
    )
    {
        XLRangeAddress xlSourceRangeAddress = (XLRangeAddress)sourceRangeAddress;
        XLRangeAddress xlTargetRangeAddress = (XLRangeAddress)targetRangeAddress;

        return this.Relative(in xlSourceRangeAddress, in xlTargetRangeAddress);
    }

    internal XLRangeAddress Relative(
        in XLRangeAddress sourceRangeAddress,
        in XLRangeAddress targetRangeAddress
    )
    {
        XLWorksheet? sheet = targetRangeAddress.Worksheet;

        return new XLRangeAddress(
            new XLAddress(
                sheet,
                this.FirstAddress.RowNumber
                    - sourceRangeAddress.FirstAddress.RowNumber
                    + targetRangeAddress.FirstAddress.RowNumber,
                this.FirstAddress.ColumnNumber
                    - sourceRangeAddress.FirstAddress.ColumnNumber
                    + targetRangeAddress.FirstAddress.ColumnNumber,
                fixedRow: false,
                fixedColumn: false
            ),
            new XLAddress(
                sheet,
                this.LastAddress.RowNumber
                    - sourceRangeAddress.FirstAddress.RowNumber
                    + targetRangeAddress.FirstAddress.RowNumber,
                this.LastAddress.ColumnNumber
                    - sourceRangeAddress.FirstAddress.ColumnNumber
                    + targetRangeAddress.FirstAddress.ColumnNumber,
                fixedRow: false,
                fixedColumn: false
            )
        );
    }

    public IXLRangeAddress Intersection(IXLRangeAddress otherRangeAddress)
    {
        ArgumentNullException.ThrowIfNull(otherRangeAddress);

        XLRangeAddress xlOtherRangeAddress = (XLRangeAddress)otherRangeAddress;
        return this.Intersection(in xlOtherRangeAddress);
    }

    internal XLRangeAddress Intersection(in XLRangeAddress otherRangeAddress)
    {
        if (!Equals(this.Worksheet, otherRangeAddress.Worksheet))
        {
            throw new ArgumentOutOfRangeException(
                nameof(otherRangeAddress),
                "The other range address is on a different worksheet"
            );
        }

        XLRangeAddress thisRangeAddressNormalized = this.Normalize();
        XLRangeAddress otherRangeAddressNormalized = otherRangeAddress.Normalize();

        int firstRow = Math.Max(
            thisRangeAddressNormalized.FirstAddress.RowNumber,
            otherRangeAddressNormalized.FirstAddress.RowNumber
        );
        int firstColumn = Math.Max(
            thisRangeAddressNormalized.FirstAddress.ColumnNumber,
            otherRangeAddressNormalized.FirstAddress.ColumnNumber
        );
        int lastRow = Math.Min(
            thisRangeAddressNormalized.LastAddress.RowNumber,
            otherRangeAddressNormalized.LastAddress.RowNumber
        );
        int lastColumn = Math.Min(
            thisRangeAddressNormalized.LastAddress.ColumnNumber,
            otherRangeAddressNormalized.LastAddress.ColumnNumber
        );

        if (lastRow < firstRow || lastColumn < firstColumn)
        {
            return XLRangeAddress.Invalid;
        }

        return new XLRangeAddress(
            new XLAddress(
                this.Worksheet,
                firstRow,
                firstColumn,
                fixedRow: false,
                fixedColumn: false
            ),
            new XLAddress(this.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
        );
    }

    public IXLRange? AsRange()
    {
        if (this.Worksheet == null)
        {
            throw new InvalidOperationException(
                "The worksheet of the current range address has not been set."
            );
        }

        if (!this.IsValid)
        {
            return null;
        }

        return this.Worksheet.Range(this);
    }

    #endregion Public methods

    #region Operators

    public static bool operator ==(XLRangeAddress left, XLRangeAddress right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(XLRangeAddress left, XLRangeAddress right)
    {
        return !(left == right);
    }

    #endregion Operators
}

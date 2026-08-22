using System;
using XlsxSharp.Excel.Rows;
using XlsxSharp.Excel.Tables;

namespace XlsxSharp.Excel;

internal class XLRangeFactory
{
    private readonly XLWorksheet _worksheet;

    public XLRangeFactory(XLWorksheet worksheet)
    {
        this._worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    #region Methods

    public XLRangeBase Create(XLRangeKey key)
    {
        switch (key.RangeType)
        {
            case XLRangeType.Range:
                return this.CreateRange(key.RangeAddress);

            case XLRangeType.Column:
                return this.CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber);

            case XLRangeType.Row:
                return this.CreateColumn(key.RangeAddress.FirstAddress.RowNumber);

            case XLRangeType.RangeColumn:
                return CreateRangeColumn(key.RangeAddress);

            case XLRangeType.RangeRow:
                return CreateRangeRow(key.RangeAddress);

            case XLRangeType.Table:
                return this.CreateTable(key.RangeAddress);

            case XLRangeType.Worksheet:
            default:
                throw new NotImplementedException(key.RangeType.ToString());
        }
    }

    public XLRange CreateRange(XLRangeAddress rangeAddress)
    {
        return new XLRange(rangeAddress, this._worksheet.Style);
    }

    public XLColumn CreateColumn(int columnNumber)
    {
        return new XLColumn(this._worksheet, columnNumber);
    }

    public XLRow CreateRow(int rowNumber)
    {
        return new XLRow(this._worksheet, rowNumber);
    }

    public static XLRangeColumn CreateRangeColumn(XLRangeAddress rangeAddress)
    {
        return new XLRangeColumn(rangeAddress);
    }

    public static XLRangeRow CreateRangeRow(XLRangeAddress rangeAddress)
    {
        return new XLRangeRow(rangeAddress);
    }

    public XLTable CreateTable(XLRangeAddress rangeAddress)
    {
        return new XLTable(rangeAddress, this._worksheet.Style);
    }

    #endregion Methods
}

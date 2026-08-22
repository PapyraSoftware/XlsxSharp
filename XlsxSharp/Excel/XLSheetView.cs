#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace XlsxSharp.Excel;

internal class XLSheetView : IXLSheetView
{
    private XLAddress _topLeftCellAddress;
    private int _zoomScale;

    public XLSheetView(XLWorksheet worksheet)
    {
        this.Worksheet = worksheet;
        this.View = XLSheetViewOptions.Normal;

        this.ZoomScale = 100;
        this.ZoomScaleNormal = 100;
        this.ZoomScalePageLayoutView = 100;
        this.ZoomScaleSheetLayoutView = 100;
    }

    public XLSheetView(XLWorksheet worksheet, XLSheetView sheetView)
        : this(worksheet)
    {
        this.SplitRow = sheetView.SplitRow;
        this.SplitColumn = sheetView.SplitColumn;
        this.FreezePanes = sheetView.FreezePanes;
        this.TopLeftCellAddress = new XLAddress(
            this.Worksheet,
            sheetView.TopLeftCellAddress.RowNumber,
            sheetView.TopLeftCellAddress.ColumnNumber,
            sheetView.TopLeftCellAddress.FixedRow,
            sheetView.TopLeftCellAddress.FixedColumn
        );
    }

    public Boolean FreezePanes { get; set; }
    public Int32 SplitColumn { get; set; }
    public Int32 SplitRow { get; set; }

    IXLAddress IXLSheetView.TopLeftCellAddress
    {
        get => this.TopLeftCellAddress;
        set => this.TopLeftCellAddress = (XLAddress)value;
    }

    public XLAddress TopLeftCellAddress
    {
        get => this._topLeftCellAddress;
        set
        {
            if (value.HasWorksheet && !value.Worksheet.Equals(this.Worksheet))
            {
                throw new ArgumentException(
                    $"The value should be on the same worksheet as the sheet view."
                );
            }

            this._topLeftCellAddress = value;
        }
    }

    public XLSheetViewOptions View { get; set; }

    IXLWorksheet IXLSheetView.Worksheet => this.Worksheet;
    public XLWorksheet Worksheet { get; internal set; }

    public int ZoomScale
    {
        get => this._zoomScale;
        set
        {
            this._zoomScale = value;
            switch (this.View)
            {
                case XLSheetViewOptions.Normal:
                    this.ZoomScaleNormal = value;
                    break;

                case XLSheetViewOptions.PageBreakPreview:
                    this.ZoomScalePageLayoutView = value;
                    break;

                case XLSheetViewOptions.PageLayout:
                    this.ZoomScaleSheetLayoutView = value;
                    break;
            }
        }
    }

    public int ZoomScaleNormal { get; set; }

    public int ZoomScalePageLayoutView { get; set; }

    public int ZoomScaleSheetLayoutView { get; set; }

    public void Freeze(Int32 rows, Int32 columns)
    {
        this.SplitRow = rows;
        this.SplitColumn = columns;
        this.FreezePanes = true;
    }

    public void FreezeColumns(Int32 columns)
    {
        this.SplitColumn = columns;
        this.FreezePanes = true;
    }

    public void FreezeRows(Int32 rows)
    {
        this.SplitRow = rows;
        this.FreezePanes = true;
    }

    public IXLSheetView SetView(XLSheetViewOptions value)
    {
        this.View = value;
        return this;
    }
}

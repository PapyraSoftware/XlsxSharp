#nullable disable

using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel.PageSetup;

internal class XLPageSetup : IXLPageSetup
{
    public XLPageSetup(XLPageSetup defaultPageOptions, XLWorksheet worksheet)
    {
        if (defaultPageOptions != null)
        {
            this.PrintAreas = new XLPrintAreas(
                defaultPageOptions.PrintAreas as XLPrintAreas,
                worksheet
            );
            this.CenterHorizontally = defaultPageOptions.CenterHorizontally;
            this.CenterVertically = defaultPageOptions.CenterVertically;
            this.FirstPageNumber = defaultPageOptions.FirstPageNumber;
            this.HorizontalDpi = defaultPageOptions.HorizontalDpi;
            this.PageOrientation = defaultPageOptions.PageOrientation;
            this.VerticalDpi = defaultPageOptions.VerticalDpi;
            this.FirstRowToRepeatAtTop = defaultPageOptions.FirstRowToRepeatAtTop;
            this.LastRowToRepeatAtTop = defaultPageOptions.LastRowToRepeatAtTop;
            this.FirstColumnToRepeatAtLeft = defaultPageOptions.FirstColumnToRepeatAtLeft;
            this.LastColumnToRepeatAtLeft = defaultPageOptions.LastColumnToRepeatAtLeft;
            this.ShowComments = defaultPageOptions.ShowComments;

            this.PaperSize = defaultPageOptions.PaperSize;
            this._pagesTall = defaultPageOptions.PagesTall;
            this._pagesWide = defaultPageOptions.PagesWide;
            this._scale = defaultPageOptions.Scale;

            if (defaultPageOptions.Margins != null)
            {
                this.Margins = new XLMargins
                {
                    Top = defaultPageOptions.Margins.Top,
                    Bottom = defaultPageOptions.Margins.Bottom,
                    Left = defaultPageOptions.Margins.Left,
                    Right = defaultPageOptions.Margins.Right,
                    Header = defaultPageOptions.Margins.Header,
                    Footer = defaultPageOptions.Margins.Footer,
                };
            }
            this.AlignHFWithMargins = defaultPageOptions.AlignHFWithMargins;
            this.ScaleHFWithDocument = defaultPageOptions.ScaleHFWithDocument;
            this.ShowGridlines = defaultPageOptions.ShowGridlines;
            this.ShowRowAndColumnHeadings = defaultPageOptions.ShowRowAndColumnHeadings;
            this.BlackAndWhite = defaultPageOptions.BlackAndWhite;
            this.DraftQuality = defaultPageOptions.DraftQuality;
            this.PageOrder = defaultPageOptions.PageOrder;

            this.ColumnBreaks = [.. defaultPageOptions.ColumnBreaks];
            this.RowBreaks = [.. defaultPageOptions.RowBreaks];
            this.Header = new XLHeaderFooter(
                defaultPageOptions.Header as XLHeaderFooter,
                worksheet
            );
            this.Footer = new XLHeaderFooter(
                defaultPageOptions.Footer as XLHeaderFooter,
                worksheet
            );
            this.PrintErrorValue = defaultPageOptions.PrintErrorValue;
        }
        else
        {
            this.PrintAreas = new XLPrintAreas(worksheet);
            this.Header = new XLHeaderFooter(worksheet);
            this.Footer = new XLHeaderFooter(worksheet);
            this.ColumnBreaks = [];
            this.RowBreaks = [];
        }
    }

    public IXLPrintAreas PrintAreas { get; private set; }

    public Int32 FirstRowToRepeatAtTop { get; private set; }
    public Int32 LastRowToRepeatAtTop { get; private set; }

    public void SetRowsToRepeatAtTop(String range)
    {
        string[] arrRange = range.Replace("$", "").Split(':');
        this.SetRowsToRepeatAtTop(Int32.Parse(arrRange[0]), Int32.Parse(arrRange[1]));
    }

    public void SetRowsToRepeatAtTop(Int32 firstRowToRepeatAtTop, Int32 lastRowToRepeatAtTop)
    {
        if (firstRowToRepeatAtTop <= 0)
        {
            throw new ArgumentOutOfRangeException("The first row has to be greater than zero.");
        }

        if (firstRowToRepeatAtTop > lastRowToRepeatAtTop)
        {
            throw new ArgumentOutOfRangeException(
                "The first row has to be less than the second row."
            );
        }

        this.FirstRowToRepeatAtTop = firstRowToRepeatAtTop;
        this.LastRowToRepeatAtTop = lastRowToRepeatAtTop;
    }

    public Int32 FirstColumnToRepeatAtLeft { get; private set; }
    public Int32 LastColumnToRepeatAtLeft { get; private set; }

    public void SetColumnsToRepeatAtLeft(String range)
    {
        string[] arrRange = range.Replace("$", "").Split(':');
        if (Int32.TryParse(arrRange[0], out int iTest))
        {
            this.SetColumnsToRepeatAtLeft(Int32.Parse(arrRange[0]), Int32.Parse(arrRange[1]));
        }
        else
        {
            this.SetColumnsToRepeatAtLeft(arrRange[0], arrRange[1]);
        }
    }

    public void SetColumnsToRepeatAtLeft(
        String firstColumnToRepeatAtLeft,
        String lastColumnToRepeatAtLeft
    )
    {
        this.SetColumnsToRepeatAtLeft(
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(firstColumnToRepeatAtLeft),
            XlsxSharp.XLHelper.GetColumnNumberFromLetter(lastColumnToRepeatAtLeft)
        );
    }

    public void SetColumnsToRepeatAtLeft(
        Int32 firstColumnToRepeatAtLeft,
        Int32 lastColumnToRepeatAtLeft
    )
    {
        if (firstColumnToRepeatAtLeft <= 0)
        {
            throw new ArgumentOutOfRangeException("The first column has to be greater than zero.");
        }

        if (firstColumnToRepeatAtLeft > lastColumnToRepeatAtLeft)
        {
            throw new ArgumentOutOfRangeException(
                "The first column has to be less than the second column."
            );
        }

        this.FirstColumnToRepeatAtLeft = firstColumnToRepeatAtLeft;
        this.LastColumnToRepeatAtLeft = lastColumnToRepeatAtLeft;
    }

    public XLPageOrientation PageOrientation { get; set; }
    public XLPaperSize PaperSize { get; set; }
    public Int32 HorizontalDpi { get; set; }
    public Int32 VerticalDpi { get; set; }
    public Int32? FirstPageNumber { get; set; }
    public Boolean CenterHorizontally { get; set; }
    public Boolean CenterVertically { get; set; }
    public XLPrintErrorValues PrintErrorValue { get; set; }
    public IXLMargins Margins { get; set; }

    private Int32 _pagesWide;
    public Int32 PagesWide
    {
        get { return this._pagesWide; }
        set
        {
            this._pagesWide = value;
            if (this._pagesWide > 0)
            {
                this._scale = 0;
            }
        }
    }

    private Int32 _pagesTall;
    public Int32 PagesTall
    {
        get { return this._pagesTall; }
        set
        {
            this._pagesTall = value;
            if (this._pagesTall > 0)
            {
                this._scale = 0;
            }
        }
    }

    private Int32 _scale;
    public Int32 Scale
    {
        get { return this._scale; }
        set
        {
            this._scale = value;
            if (this._scale <= 0)
            {
                return;
            }

            this._pagesTall = 0;
            this._pagesWide = 0;
        }
    }

    public void AdjustTo(Int32 percentageOfNormalSize)
    {
        this.Scale = percentageOfNormalSize;
        this._pagesWide = 0;
        this._pagesTall = 0;
    }

    public void FitToPages(Int32 pagesWide, Int32 pagesTall)
    {
        this._pagesWide = pagesWide;
        this._pagesTall = pagesTall;
        this._scale = 0;
    }

    public IXLHeaderFooter Header { get; private set; }
    public IXLHeaderFooter Footer { get; private set; }

    public Boolean ScaleHFWithDocument { get; set; }
    public Boolean AlignHFWithMargins { get; set; }

    public Boolean ShowGridlines { get; set; }
    public Boolean ShowRowAndColumnHeadings { get; set; }
    public Boolean BlackAndWhite { get; set; }
    public Boolean DraftQuality { get; set; }

    public XLPageOrderValues PageOrder { get; set; }
    public XLShowCommentsValues ShowComments { get; set; }

    public List<Int32> RowBreaks { get; private set; }
    public List<Int32> ColumnBreaks { get; private set; }

    public void AddHorizontalPageBreak(Int32 row)
    {
        if (!this.RowBreaks.Contains(row))
        {
            this.RowBreaks.Add(row);
        }

        this.RowBreaks.Sort();
    }

    public void AddVerticalPageBreak(Int32 column)
    {
        if (!this.ColumnBreaks.Contains(column))
        {
            this.ColumnBreaks.Add(column);
        }

        this.ColumnBreaks.Sort();
    }

    //public void SetPageBreak(IXLRange range, XLPageBreakLocations breakLocation)
    //{
    //    switch (breakLocation)
    //    {
    //        case XLPageBreakLocations.AboveRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowNumber)); break;
    //        case XLPageBreakLocations.BelowRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowCount())); break;
    //        case XLPageBreakLocations.LeftOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnNumber)); break;
    //        case XLPageBreakLocations.RightOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnCount())); break;
    //        default: throw new NotImplementedException();
    //    }
    //}

    public IXLPageSetup SetPageOrientation(XLPageOrientation value)
    {
        this.PageOrientation = value;
        return this;
    }

    public IXLPageSetup SetPagesWide(Int32 value)
    {
        this.PagesWide = value;
        return this;
    }

    public IXLPageSetup SetPagesTall(Int32 value)
    {
        this.PagesTall = value;
        return this;
    }

    public IXLPageSetup SetScale(Int32 value)
    {
        this.Scale = value;
        return this;
    }

    public IXLPageSetup SetHorizontalDpi(Int32 value)
    {
        this.HorizontalDpi = value;
        return this;
    }

    public IXLPageSetup SetVerticalDpi(Int32 value)
    {
        this.VerticalDpi = value;
        return this;
    }

    public IXLPageSetup SetFirstPageNumber(Int32? value)
    {
        this.FirstPageNumber = value;
        return this;
    }

    public IXLPageSetup SetCenterHorizontally()
    {
        this.CenterHorizontally = true;
        return this;
    }

    public IXLPageSetup SetCenterHorizontally(Boolean value)
    {
        this.CenterHorizontally = value;
        return this;
    }

    public IXLPageSetup SetCenterVertically()
    {
        this.CenterVertically = true;
        return this;
    }

    public IXLPageSetup SetCenterVertically(Boolean value)
    {
        this.CenterVertically = value;
        return this;
    }

    public IXLPageSetup SetPaperSize(XLPaperSize value)
    {
        this.PaperSize = value;
        return this;
    }

    public IXLPageSetup SetScaleHFWithDocument()
    {
        this.ScaleHFWithDocument = true;
        return this;
    }

    public IXLPageSetup SetScaleHFWithDocument(Boolean value)
    {
        this.ScaleHFWithDocument = value;
        return this;
    }

    public IXLPageSetup SetAlignHFWithMargins()
    {
        this.AlignHFWithMargins = true;
        return this;
    }

    public IXLPageSetup SetAlignHFWithMargins(Boolean value)
    {
        this.AlignHFWithMargins = value;
        return this;
    }

    public IXLPageSetup SetShowGridlines()
    {
        this.ShowGridlines = true;
        return this;
    }

    public IXLPageSetup SetShowGridlines(Boolean value)
    {
        this.ShowGridlines = value;
        return this;
    }

    public IXLPageSetup SetShowRowAndColumnHeadings()
    {
        this.ShowRowAndColumnHeadings = true;
        return this;
    }

    public IXLPageSetup SetShowRowAndColumnHeadings(Boolean value)
    {
        this.ShowRowAndColumnHeadings = value;
        return this;
    }

    public IXLPageSetup SetBlackAndWhite()
    {
        this.BlackAndWhite = true;
        return this;
    }

    public IXLPageSetup SetBlackAndWhite(Boolean value)
    {
        this.BlackAndWhite = value;
        return this;
    }

    public IXLPageSetup SetDraftQuality()
    {
        this.DraftQuality = true;
        return this;
    }

    public IXLPageSetup SetDraftQuality(Boolean value)
    {
        this.DraftQuality = value;
        return this;
    }

    public IXLPageSetup SetPageOrder(XLPageOrderValues value)
    {
        this.PageOrder = value;
        return this;
    }

    public IXLPageSetup SetShowComments(XLShowCommentsValues value)
    {
        this.ShowComments = value;
        return this;
    }

    public IXLPageSetup SetPrintErrorValue(XLPrintErrorValues value)
    {
        this.PrintErrorValue = value;
        return this;
    }

    public Boolean DifferentFirstPageOnHF { get; set; }

    public IXLPageSetup SetDifferentFirstPageOnHF()
    {
        return this.SetDifferentFirstPageOnHF(true);
    }

    public IXLPageSetup SetDifferentFirstPageOnHF(Boolean value)
    {
        this.DifferentFirstPageOnHF = value;
        return this;
    }

    public Boolean DifferentOddEvenPagesOnHF { get; set; }

    public IXLPageSetup SetDifferentOddEvenPagesOnHF()
    {
        return this.SetDifferentOddEvenPagesOnHF(true);
    }

    public IXLPageSetup SetDifferentOddEvenPagesOnHF(Boolean value)
    {
        this.DifferentOddEvenPagesOnHF = value;
        return this;
    }
}

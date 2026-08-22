#nullable disable warnings

using System;
using System.Diagnostics;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Excel.Drawings.Style;
using XlsxSharp.Excel.Formatting;
using XlsxSharp.Excel.RichText;
using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Excel.Comments;

internal class XLComment : XLFormattedText<IXLComment>, IXLComment
{
    private XLCell _cell;

    private static XLFontFormatValue DefaultCommentFont
    {
        get
        {
            // MS Excel uses Tahoma 9 Swiss no matter what current style font
            XLFontFormatValue? defaultCommentFont = XLFontFormatValue.Default with
            {
                Name = "Tahoma",
                Size = XLFontSize.FromPoints(9),
                Family = XLFontFamilyNumberingValues.Swiss,
                Color = XLColor.Black,
            };

            return defaultCommentFont;
        }
    }

    private XLComment(
        XLFontFormatValue defaultFont,
        XLWorkbookStyles styles,
        XLFontFormatValue? phoneticsFont
    )
        : base(defaultFont, styles)
    {
        if (phoneticsFont is not null)
        {
            Debug.Assert(styles.Fonts.ContainsValue(phoneticsFont));
            this.Phonetics = new XLPhonetics(
                phoneticsFont,
                defaultFont,
                styles,
                this.OnContentChanged
            );
        }
    }

    #region IXLComment Members

    public String Author { get; set; }

    public IXLComment SetAuthor(String value)
    {
        this.Author = value;
        return this;
    }

    public IXLRichString AddSignature()
    {
        this.AddText(this.Author + ":").SetBold();
        return this.AddText(Environment.NewLine);
    }

    public void Delete() => this._cell.DeleteComment();

    #endregion IXLComment Members

    #region IXLDrawing

    public String Name { get; set; }
    public String Description { get; set; }
    public XLDrawingAnchor Anchor { get; set; }
    public Boolean HorizontalFlip { get; set; }
    public Boolean VerticalFlip { get; set; }
    public Int32 Rotation { get; set; }
    public Int32 ExtentLength { get; set; }
    public Int32 ExtentWidth { get; set; }
    public Int32 ShapeId { get; internal set; }
    public Boolean Visible { get; set; }

    public IXLComment SetVisible()
    {
        this.Visible = true;
        return this.Container;
    }

    public IXLComment SetVisible(Boolean hidden)
    {
        this.Visible = hidden;
        return this.Container;
    }

    public IXLDrawingPosition Position { get; private set; }

    public Int32 ZOrder { get; set; }

    public IXLComment SetZOrder(Int32 zOrder)
    {
        this.ZOrder = zOrder;
        return this.Container;
    }

    public IXLDrawingStyle Style { get; private set; }

    public IXLComment SetName(String name)
    {
        this.Name = name;
        return this.Container;
    }

    public IXLComment SetDescription(String description)
    {
        this.Description = description;
        return this.Container;
    }

    public IXLComment SetHorizontalFlip()
    {
        this.HorizontalFlip = true;
        return this.Container;
    }

    public IXLComment SetHorizontalFlip(Boolean horizontalFlip)
    {
        this.HorizontalFlip = horizontalFlip;
        return this.Container;
    }

    public IXLComment SetVerticalFlip()
    {
        this.VerticalFlip = true;
        return this.Container;
    }

    public IXLComment SetVerticalFlip(Boolean verticalFlip)
    {
        this.VerticalFlip = verticalFlip;
        return this.Container;
    }

    public IXLComment SetRotation(Int32 rotation)
    {
        this.Rotation = rotation;
        return this.Container;
    }

    public IXLComment SetExtentLength(Int32 extentLength)
    {
        this.ExtentLength = extentLength;
        return this.Container;
    }

    public IXLComment SetExtentWidth(Int32 extentWidth)
    {
        this.ExtentWidth = extentWidth;
        return this.Container;
    }

    #endregion IXLDrawing

    internal static XLComment Create(XLCell cell, int? shapeId)
    {
        XLWorkbookStyles? styles = cell.Worksheet.Workbook.Styles;
        XLFontFormatValue? defaultFont = styles.RegisterFontFormat(DefaultCommentFont);
        XLComment? comment = new(defaultFont, styles, null);
        comment.Initialize(cell, shapeId: shapeId);
        return comment;
    }

    internal static XLComment CreateAsCopy(
        XLCell targetCell,
        XLCell sourceCell,
        XLComment originalComment
    )
    {
        // source cell could be from different workbook, so register formats
        XLWorkbookStyles? styles = targetCell.Worksheet.Workbook.Styles;
        XLFontFormatValue? defaultFont = styles.RegisterFontFormat(sourceCell.GetFormat().Font);
        XLFontFormatValue? phoneticsFont = originalComment.HasPhonetics
            ? styles.RegisterFontFormat(
                XLFontFormatValue.FromFontBase(originalComment.Phonetics, styles)
            )
            : null;
        XLComment? comment = new(defaultFont, styles, phoneticsFont);

        foreach (XLRichString rt in originalComment)
        {
            comment.AddText(rt.Text, rt);
        }

        comment.Initialize(targetCell, originalComment.Style);
        return comment;
    }

    private void Initialize(XLCell cell, IXLDrawingStyle style = null, int? shapeId = null)
    {
        style = style ?? XLDrawingStyle.DefaultCommentStyle;
        shapeId = shapeId ?? cell.Worksheet.Workbook.ShapeIdManager.GetNext();

        this.Author = cell.Worksheet.Author;
        this.Container = this;
        this.Anchor = XLDrawingAnchor.MoveAndSizeWithCells;
        this.Style = new XLDrawingStyle();
        Int32 previousRowNumber = cell.Address.RowNumber;
        Double previousRowOffset = 0;

        if (previousRowNumber > 1)
        {
            previousRowNumber--;

            if (
                cell.Worksheet.Internals.RowsCollection.TryGetValue(
                    previousRowNumber,
                    out XLRow previousRow
                )
            )
            {
                previousRowOffset = Math.Max(0, previousRow.Height - 7);
            }
            else
            {
                previousRowOffset = Math.Max(0, cell.Worksheet.RowHeight - 7);
            }
        }

        this.Position = new XLDrawingPosition
        {
            Column = cell.Address.ColumnNumber + 1,
            ColumnOffset = 2,
            Row = previousRowNumber,
            RowOffset = previousRowOffset,
        };

        this.ZOrder = cell.Worksheet.ZOrder++;
        this.Style.Margins.SetLeft(style.Margins.Left)
            .Margins.SetRight(style.Margins.Right)
            .Margins.SetTop(style.Margins.Top)
            .Margins.SetBottom(style.Margins.Bottom)
            .Margins.SetAutomatic(style.Margins.Automatic)
            .Size.SetHeight(style.Size.Height)
            .Size.SetWidth(style.Size.Width)
            .ColorsAndLines.SetLineColor(style.ColorsAndLines.LineColor)
            .ColorsAndLines.SetFillColor(style.ColorsAndLines.FillColor)
            .ColorsAndLines.SetLineDash(style.ColorsAndLines.LineDash)
            .ColorsAndLines.SetLineStyle(style.ColorsAndLines.LineStyle)
            .ColorsAndLines.SetLineWeight(style.ColorsAndLines.LineWeight)
            .ColorsAndLines.SetFillTransparency(style.ColorsAndLines.FillTransparency)
            .ColorsAndLines.SetLineTransparency(style.ColorsAndLines.LineTransparency)
            .Alignment.SetHorizontal(style.Alignment.Horizontal)
            .Alignment.SetVertical(style.Alignment.Vertical)
            .Alignment.SetDirection(style.Alignment.Direction)
            .Alignment.SetOrientation(style.Alignment.Orientation)
            .Alignment.SetAutomaticSize(style.Alignment.AutomaticSize)
            .Properties.SetPositioning(style.Properties.Positioning)
            .Protection.SetLocked(style.Protection.Locked)
            .Protection.SetLockText(style.Protection.LockText);

        this._cell = cell;
        this.ShapeId = shapeId.Value;
    }
}

#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using XlsxSharp.Graphics;

namespace XlsxSharp.Excel.Drawings;

[DebuggerDisplay("{Name}")]
internal sealed class XLPicture : IXLPicture, IDisposable
{
    private Int32 _height;
    private Int32 _id;
    private String _name = string.Empty;
    private Int32 _width;
    private bool _disposed;

    internal XLPicture(XLWorksheet worksheet, Stream stream)
        : this(worksheet, stream, XLPictureFormat.Unknown) { }

    internal XLPicture(IXLWorksheet worksheet, Stream stream, XLPictureFormat format)
        : this(worksheet)
    {
        ArgumentNullException.ThrowIfNull(stream);

        XLPictureInfo info = worksheet.Workbook.GraphicEngine.GetPictureInfo(stream, format);
        this.Init(info);

        this.ImageStream = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(this.ImageStream);
        this.ImageStream.Seek(0, SeekOrigin.Begin);
    }

    private XLPicture(IXLWorksheet worksheet)
    {
        this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        this.Placement = XLPicturePlacement.MoveAndSize;
        this.Markers = new Dictionary<XLMarkerPosition, XLMarker>()
        {
            [XLMarkerPosition.TopLeft] = null,
            [XLMarkerPosition.BottomRight] = null,
        };

        // Calculate default picture ID
        IEnumerable<IXLPicture> allPictures = worksheet.Workbook.Worksheets.SelectMany(ws =>
            ws.Pictures
        );
        int freeId = allPictures.Select(x => x.Id).DefaultIfEmpty(0).Max() + 1;
        this._id = freeId;
    }

    public IXLCell BottomRightCell
    {
        get => this.Markers[XLMarkerPosition.BottomRight].Cell;
        private set
        {
            if (!value.Worksheet.Equals(this.Worksheet))
            {
                throw new InvalidOperationException(
                    "A picture and its anchor cells must be on the same worksheet"
                );
            }

            this.Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
        }
    }

    public XLPictureFormat Format { get; private set; } = XLPictureFormat.Unknown;

    public Int32 Height
    {
        get => this._height;
        set
        {
            if (this.Placement == XLPicturePlacement.MoveAndSize)
            {
                throw new ArgumentException(
                    "To set the height, the placement should be FreeFloating or Move"
                );
            }

            this._height = value;
        }
    }

    public Int32 Id
    {
        get => this._id;
        internal set
        {
            if ((this.Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
            {
                throw new ArgumentException($"The picture ID '{value}' already exists.");
            }

            this._id = value;
        }
    }

    public MemoryStream ImageStream
    {
        get
        {
            this.ThrowIfDisposed();
            return field;
        }
        private init;
    }

    public Int32 Left
    {
        get => this.Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0;
        set
        {
            if (this.Placement != XLPicturePlacement.FreeFloating)
            {
                throw new ArgumentException(
                    "To set the left-hand offset, the placement should be FreeFloating"
                );
            }

            this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(
                this.Worksheet.Cell(1, 1),
                new System.Drawing.Point(value, this.Top)
            );
        }
    }

    public String Name
    {
        get => this._name;
        set
        {
            if (this._name == value)
            {
                return;
            }

            if (
                (
                    this.Worksheet.Pictures.FirstOrDefault(p =>
                        p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)
                    ) ?? this
                ) != this
            )
            {
                throw new ArgumentException($"The picture name '{value}' already exists.");
            }

            this.SetName(value);
        }
    }

    public Int32 OriginalHeight { get; private set; }

    public Int32 OriginalWidth { get; private set; }

    public XLPicturePlacement Placement { get; set; }

    public Int32 Top
    {
        get => this.Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0;
        set
        {
            if (this.Placement != XLPicturePlacement.FreeFloating)
            {
                throw new ArgumentException(
                    "To set the top offset, the placement should be FreeFloating"
                );
            }

            this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(
                this.Worksheet.Cell(1, 1),
                new System.Drawing.Point(this.Left, value)
            );
        }
    }

    public IXLCell TopLeftCell
    {
        get => this.Markers[XLMarkerPosition.TopLeft].Cell;
        private set
        {
            if (!value.Worksheet.Equals(this.Worksheet))
            {
                throw new InvalidOperationException(
                    "A picture and its anchor cells must be on the same worksheet"
                );
            }

            this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
        }
    }

    public Int32 Width
    {
        get => this._width;
        set
        {
            if (this.Placement == XLPicturePlacement.MoveAndSize)
            {
                throw new ArgumentException(
                    "To set the width, the placement should be FreeFloating or Move"
                );
            }

            this._width = value;
        }
    }

    public IXLWorksheet Worksheet { get; }

    internal IDictionary<XLMarkerPosition, XLMarker> Markers { get; private set; }

    internal String RelId { get; set; }

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture CopyTo(IXLWorksheet targetSheet) => this.CopyTo((XLWorksheet)targetSheet);

    public void Delete() => this.Worksheet.Pictures.Delete(this.Name);

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture Duplicate() => this.CopyTo(this.Worksheet);

    public System.Drawing.Point GetOffset(XLMarkerPosition position) =>
        this.Markers[position].Offset;

    public IXLPicture MoveTo(Int32 left, Int32 top)
    {
        this.Placement = XLPicturePlacement.FreeFloating;
        this.Left = left;
        this.Top = top;
        return this;
    }

    public IXLPicture MoveTo(IXLCell cell) => this.MoveTo(cell, 0, 0);

    public IXLPicture MoveTo(IXLCell cell, Int32 xOffset, Int32 yOffset) =>
        this.MoveTo(cell, new System.Drawing.Point(xOffset, yOffset));

    public IXLPicture MoveTo(IXLCell cell, System.Drawing.Point offset)
    {
        ArgumentNullException.ThrowIfNull(cell);
        this.Placement = XLPicturePlacement.Move;
        this.TopLeftCell = cell;
        this.Markers[XLMarkerPosition.TopLeft].Offset = offset;
        return this;
    }

    public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell) =>
        this.MoveTo(fromCell, 0, 0, toCell, 0, 0);

    public IXLPicture MoveTo(
        IXLCell fromCell,
        Int32 fromCellXOffset,
        Int32 fromCellYOffset,
        IXLCell toCell,
        Int32 toCellXOffset,
        Int32 toCellYOffset
    ) =>
        this.MoveTo(
            fromCell,
            new System.Drawing.Point(fromCellXOffset, fromCellYOffset),
            toCell,
            new System.Drawing.Point(toCellXOffset, toCellYOffset)
        );

    public IXLPicture MoveTo(
        IXLCell fromCell,
        System.Drawing.Point fromOffset,
        IXLCell toCell,
        System.Drawing.Point toOffset
    )
    {
        ArgumentNullException.ThrowIfNull(fromCell);
        ArgumentNullException.ThrowIfNull(toCell);
        this.Placement = XLPicturePlacement.MoveAndSize;

        this.TopLeftCell = fromCell;
        this.Markers[XLMarkerPosition.TopLeft].Offset = fromOffset;

        this.BottomRightCell = toCell;
        this.Markers[XLMarkerPosition.BottomRight].Offset = toOffset;

        return this;
    }

    public IXLPicture Scale(Double factor, Boolean relativeToOriginal = false) =>
        this.ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);

    public IXLPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false)
    {
        this.Height = Convert.ToInt32(
            (relativeToOriginal ? this.OriginalHeight : this.Height) * factor
        );
        return this;
    }

    public IXLPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false)
    {
        this.Width = Convert.ToInt32(
            (relativeToOriginal ? this.OriginalWidth : this.Width) * factor
        );
        return this;
    }

    public IXLPicture WithPlacement(XLPicturePlacement value)
    {
        this.Placement = value;
        return this;
    }

    public IXLPicture WithSize(Int32 width, Int32 height)
    {
        this.Width = width;
        this.Height = height;
        return this;
    }

    public void Dispose()
    {
        if (this._disposed)
        {
            return;
        }

        this.ImageStream.Dispose();
        this._disposed = true;
    }

    internal IXLPicture CopyTo(XLWorksheet targetSheet)
    {
        this.ThrowIfDisposed();
        if (targetSheet == null)
        {
            targetSheet = this.Worksheet as XLWorksheet;
        }

        IXLPicture newPicture;
        if (targetSheet == this.Worksheet)
        {
            newPicture = targetSheet.AddPicture(this.ImageStream, this.Format);
        }
        else
        {
            newPicture = targetSheet.AddPicture(this.ImageStream, this.Format, this.Name);
        }

        newPicture = newPicture
            .WithPlacement(XLPicturePlacement.FreeFloating)
            .WithSize(this.Width, this.Height)
            .WithPlacement(this.Placement);

        switch (this.Placement)
        {
            case XLPicturePlacement.FreeFloating:
                newPicture.MoveTo(this.Left, this.Top);
                break;

            case XLPicturePlacement.Move:
                newPicture.MoveTo(
                    targetSheet.Cell(this.TopLeftCell.Address),
                    this.GetOffset(XLMarkerPosition.TopLeft)
                );
                break;

            case XLPicturePlacement.MoveAndSize:
                newPicture.MoveTo(
                    targetSheet.Cell(this.TopLeftCell.Address),
                    this.GetOffset(XLMarkerPosition.TopLeft),
                    targetSheet.Cell(this.BottomRightCell.Address),
                    this.GetOffset(XLMarkerPosition.BottomRight)
                );
                break;
        }

        return newPicture;
    }

    internal void SetName(string value)
    {
        if (String.IsNullOrEmpty(value))
        {
            throw new ArgumentException("Picture names cannot be null or empty.");
        }

        this._name = value;
    }

    private void Init(XLPictureInfo info)
    {
        this.Format = info.Format;
        Size size = info.GetSizePx(this.Worksheet.Workbook.DpiX, this.Worksheet.Workbook.DpiY);
        this._width = this.OriginalWidth = size.Width;
        this._height = this.OriginalHeight = size.Height;
    }

    private void ThrowIfDisposed()
    {
        if (this._disposed)
        {
            throw new ObjectDisposedException(nameof(XLPicture));
        }
    }
}

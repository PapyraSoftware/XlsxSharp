#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.IO;

namespace XlsxSharp.Excel.Drawings;

public interface IXLPicture
{
    public IXLCell BottomRightCell { get; }

    /// <summary>
    /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
    /// Default value is "jpeg"
    /// </summary>
    public XLPictureFormat Format { get; }

    /// <summary>
    /// Current width of the picture in pixels.
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// Current height of the picture in pixels.
    /// </summary>
    public int Height { get; set; }

    public int Id { get; }

    public MemoryStream ImageStream { get; }

    public int Left { get; set; }

    /// <summary>
    /// Set the name of a picture.
    /// </summary>
    /// <exception cref="ArgumentException">Name is already used in the sheet, is null or empty.</exception>
    public string Name { get; set; }

    /// <summary>
    /// Original height of the picture in pixels.
    /// </summary>
    public int OriginalHeight { get; }

    /// <summary>
    /// Original width of the picture in pixels.
    /// </summary>
    public int OriginalWidth { get; }

    public XLPicturePlacement Placement { get; set; }

    public int Top { get; set; }

    public IXLCell TopLeftCell { get; }

    public IXLWorksheet Worksheet { get; }

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture CopyTo(IXLWorksheet targetSheet);

    /// <summary>
    /// Deletes this picture.
    /// </summary>
    public void Delete();

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    public IXLPicture Duplicate();

    public System.Drawing.Point GetOffset(XLMarkerPosition position);

    public IXLPicture MoveTo(int left, int top);

    public IXLPicture MoveTo(IXLCell cell);

    public IXLPicture MoveTo(IXLCell cell, int xOffset, int yOffset);

    public IXLPicture MoveTo(IXLCell cell, System.Drawing.Point offset);

    public IXLPicture MoveTo(IXLCell fromCell, IXLCell toCell);

    public IXLPicture MoveTo(
        IXLCell fromCell,
        int fromCellXOffset,
        int fromCellYOffset,
        IXLCell toCell,
        int toCellXOffset,
        int toCellYOffset
    );

    public IXLPicture MoveTo(
        IXLCell fromCell,
        System.Drawing.Point fromOffset,
        IXLCell toCell,
        System.Drawing.Point toOffset
    );

    public IXLPicture Scale(double factor, bool relativeToOriginal = false);

    public IXLPicture ScaleHeight(double factor, bool relativeToOriginal = false);

    public IXLPicture ScaleWidth(double factor, bool relativeToOriginal = false);

    public IXLPicture WithPlacement(XLPicturePlacement value);

    public IXLPicture WithSize(int width, int height);
}

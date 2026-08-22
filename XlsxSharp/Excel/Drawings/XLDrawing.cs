#nullable disable

using System;
using XlsxSharp.Excel.Drawings.Style;

namespace XlsxSharp.Excel.Drawings;

internal class XLDrawing<T> : IXLDrawing<T>
{
    internal T Container;

    public XLDrawing()
    {
        this.Style = new XLDrawingStyle();
        this.Position = new XLDrawingPosition();
    }

    public int ShapeId { get; internal set; }

    public bool Visible { get; set; }

    public T SetVisible()
    {
        this.Visible = true;
        return this.Container;
    }

    public T SetVisible(bool hidden)
    {
        this.Visible = hidden;
        return this.Container;
    }

    public string Name { get; set; }

    public T SetName(string name)
    {
        this.Name = name;
        return this.Container;
    }

    public string Description { get; set; }

    public T SetDescription(string description)
    {
        this.Description = description;
        return this.Container;
    }

    public IXLDrawingPosition Position { get; private set; }

    public int ZOrder { get; set; }

    public T SetZOrder(int zOrder)
    {
        this.ZOrder = zOrder;
        return this.Container;
    }

    public bool HorizontalFlip { get; set; }

    public T SetHorizontalFlip()
    {
        this.HorizontalFlip = true;
        return this.Container;
    }

    public T SetHorizontalFlip(bool horizontalFlip)
    {
        this.HorizontalFlip = horizontalFlip;
        return this.Container;
    }

    public bool VerticalFlip { get; set; }

    public T SetVerticalFlip()
    {
        this.VerticalFlip = true;
        return this.Container;
    }

    public T SetVerticalFlip(bool verticalFlip)
    {
        this.VerticalFlip = verticalFlip;
        return this.Container;
    }

    public int Rotation { get; set; }

    public T SetRotation(int rotation)
    {
        this.Rotation = rotation;
        return this.Container;
    }

    public int OffsetX { get; set; }

    public T SetOffsetX(int offsetX)
    {
        this.OffsetX = offsetX;
        return this.Container;
    }

    public int OffsetY { get; set; }

    public T SetOffsetY(int offsetY)
    {
        this.OffsetY = offsetY;
        return this.Container;
    }

    public int ExtentLength { get; set; }

    public T SetExtentLength(int extentLength)
    {
        this.ExtentLength = extentLength;
        return this.Container;
    }

    public int ExtentWidth { get; set; }

    public T SetExtentWidth(int extentWidth)
    {
        this.ExtentWidth = extentWidth;
        return this.Container;
    }

    public IXLDrawingStyle Style { get; private set; }
}

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

    public Int32 ShapeId { get; internal set; }

    public Boolean Visible { get; set; }

    public T SetVisible()
    {
        this.Visible = true;
        return this.Container;
    }

    public T SetVisible(Boolean hidden)
    {
        this.Visible = hidden;
        return this.Container;
    }

    public String Name { get; set; }

    public T SetName(String name)
    {
        this.Name = name;
        return this.Container;
    }

    public String Description { get; set; }

    public T SetDescription(String description)
    {
        this.Description = description;
        return this.Container;
    }

    public IXLDrawingPosition Position { get; private set; }

    public Int32 ZOrder { get; set; }

    public T SetZOrder(Int32 zOrder)
    {
        this.ZOrder = zOrder;
        return this.Container;
    }

    public Boolean HorizontalFlip { get; set; }

    public T SetHorizontalFlip()
    {
        this.HorizontalFlip = true;
        return this.Container;
    }

    public T SetHorizontalFlip(Boolean horizontalFlip)
    {
        this.HorizontalFlip = horizontalFlip;
        return this.Container;
    }

    public Boolean VerticalFlip { get; set; }

    public T SetVerticalFlip()
    {
        this.VerticalFlip = true;
        return this.Container;
    }

    public T SetVerticalFlip(Boolean verticalFlip)
    {
        this.VerticalFlip = verticalFlip;
        return this.Container;
    }

    public Int32 Rotation { get; set; }

    public T SetRotation(Int32 rotation)
    {
        this.Rotation = rotation;
        return this.Container;
    }

    public Int32 OffsetX { get; set; }

    public T SetOffsetX(Int32 offsetX)
    {
        this.OffsetX = offsetX;
        return this.Container;
    }

    public Int32 OffsetY { get; set; }

    public T SetOffsetY(Int32 offsetY)
    {
        this.OffsetY = offsetY;
        return this.Container;
    }

    public Int32 ExtentLength { get; set; }

    public T SetExtentLength(Int32 extentLength)
    {
        this.ExtentLength = extentLength;
        return this.Container;
    }

    public Int32 ExtentWidth { get; set; }

    public T SetExtentWidth(Int32 extentWidth)
    {
        this.ExtentWidth = extentWidth;
        return this.Container;
    }

    public IXLDrawingStyle Style { get; private set; }
}

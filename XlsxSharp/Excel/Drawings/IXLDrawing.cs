#nullable disable

using XlsxSharp.Excel.Drawings.Style;

namespace XlsxSharp.Excel.Drawings;

public enum XLDrawingAnchor
{
    MoveAndSizeWithCells,
    MoveWithCells,
    Absolute,
}

public interface IXLDrawing<T>
{
    public int ShapeId { get; }

    public bool Visible { get; set; }
    public T SetVisible();
    public T SetVisible(bool hidden);

    ////String Name { get; set; }
    ////T SetName(String name);

    ////String Description { get; set; }
    ////T SetDescription(String description);

    public IXLDrawingPosition Position { get; }

    public int ZOrder { get; set; }
    public T SetZOrder(int zOrder);

    //Boolean HorizontalFlip { get; set; }
    //T SetHorizontalFlip();
    //T SetHorizontalFlip(Boolean horizontalFlip);

    //Boolean VerticalFlip { get; set; }
    //T SetVerticalFlip();
    //T SetVerticalFlip(Boolean verticalFlip);

    //Int32 Rotation { get; set; }
    //T SetRotation(Int32 rotation);

    //Int32 ExtentLength { get; set; }
    //T SetExtentLength(Int32 ExtentLength);

    //Int32 ExtentWidth { get; set; }
    //T SetExtentWidth(Int32 extentWidth);

    public IXLDrawingStyle Style { get; }
}

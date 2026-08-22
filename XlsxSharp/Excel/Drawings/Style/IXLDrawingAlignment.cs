#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public enum XLDrawingTextDirection
{
    Context,
    LeftToRight,
    RightToLeft,
}

public enum XLDrawingTextOrientation
{
    LeftToRight,
    Vertical,
    BottomToTop,
    TopToBottom,
}

public enum XLDrawingHorizontalAlignment
{
    Left,
    Justify,
    Center,
    Right,
    Distributed,
}

public enum XLDrawingVerticalAlignment
{
    Top,
    Justify,
    Center,
    Bottom,
    Distributed,
}

public interface IXLDrawingAlignment
{
    public XLDrawingHorizontalAlignment Horizontal { get; set; }
    public XLDrawingVerticalAlignment Vertical { get; set; }
    public bool AutomaticSize { get; set; }
    public XLDrawingTextDirection Direction { get; set; }
    public XLDrawingTextOrientation Orientation { get; set; }

    public IXLDrawingStyle SetHorizontal(XLDrawingHorizontalAlignment value);
    public IXLDrawingStyle SetVertical(XLDrawingVerticalAlignment value);
    public IXLDrawingStyle SetAutomaticSize();
    public IXLDrawingStyle SetAutomaticSize(bool value);
    public IXLDrawingStyle SetDirection(XLDrawingTextDirection value);
    public IXLDrawingStyle SetOrientation(XLDrawingTextOrientation value);
}

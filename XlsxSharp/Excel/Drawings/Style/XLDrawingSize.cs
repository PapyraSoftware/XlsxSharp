#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingSize : IXLDrawingSize
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingSize(IXLDrawingStyle style) => this._style = style;

    public bool AutomaticSize
    {
        get => this._style.Alignment.AutomaticSize;
        set => this._style.Alignment.AutomaticSize = value;
    }

    public IXLDrawingStyle SetAutomaticSize()
    {
        this.AutomaticSize = true;
        return this._style;
    }

    public IXLDrawingStyle SetAutomaticSize(bool value)
    {
        this.AutomaticSize = value;
        return this._style;
    }

    public double Height { get; set; }

    public IXLDrawingStyle SetHeight(double value)
    {
        this.Height = value;
        return this._style;
    }

    public double Width { get; set; }

    public IXLDrawingStyle SetWidth(double value)
    {
        this.Width = value;
        return this._style;
    }
}

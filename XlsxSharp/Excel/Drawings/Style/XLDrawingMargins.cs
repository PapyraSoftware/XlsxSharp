#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingMargins : IXLDrawingMargins
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingMargins(IXLDrawingStyle style) => this._style = style;

    public bool Automatic { get; set; }

    public IXLDrawingStyle SetAutomatic()
    {
        this.Automatic = true;
        return this._style;
    }

    public IXLDrawingStyle SetAutomatic(bool value)
    {
        this.Automatic = value;
        return this._style;
    }

    private double _left;
    public double Left
    {
        get => this._left;
        set
        {
            this._left = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetLeft(double value)
    {
        this.Left = value;
        return this._style;
    }

    private double _right;
    public double Right
    {
        get => this._right;
        set
        {
            this._right = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetRight(double value)
    {
        this.Right = value;
        return this._style;
    }

    private double _top;
    public double Top
    {
        get => this._top;
        set
        {
            this._top = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetTop(double value)
    {
        this.Top = value;
        return this._style;
    }

    private double _bottom;
    public double Bottom
    {
        get => this._bottom;
        set
        {
            this._bottom = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetBottom(double value)
    {
        this.Bottom = value;
        return this._style;
    }

    public double All
    {
        set
        {
            this._left = value;
            this._right = value;
            this._top = value;
            this._bottom = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetAll(double value)
    {
        this.All = value;
        return this._style;
    }
}

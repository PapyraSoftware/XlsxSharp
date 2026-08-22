#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingMargins : IXLDrawingMargins
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingMargins(IXLDrawingStyle style) => this._style = style;

    public Boolean Automatic { get; set; }

    public IXLDrawingStyle SetAutomatic()
    {
        this.Automatic = true;
        return this._style;
    }

    public IXLDrawingStyle SetAutomatic(Boolean value)
    {
        this.Automatic = value;
        return this._style;
    }

    Double _left;
    public Double Left
    {
        get => this._left;
        set
        {
            this._left = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetLeft(Double value)
    {
        this.Left = value;
        return this._style;
    }

    Double _right;
    public Double Right
    {
        get => this._right;
        set
        {
            this._right = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetRight(Double value)
    {
        this.Right = value;
        return this._style;
    }

    Double _top;
    public Double Top
    {
        get => this._top;
        set
        {
            this._top = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetTop(Double value)
    {
        this.Top = value;
        return this._style;
    }

    Double _bottom;
    public Double Bottom
    {
        get => this._bottom;
        set
        {
            this._bottom = value;
            this.Automatic = false;
        }
    }

    public IXLDrawingStyle SetBottom(Double value)
    {
        this.Bottom = value;
        return this._style;
    }

    public Double All
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

    public IXLDrawingStyle SetAll(Double value)
    {
        this.All = value;
        return this._style;
    }
}

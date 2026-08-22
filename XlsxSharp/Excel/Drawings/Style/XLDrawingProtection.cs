#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingProtection : IXLDrawingProtection
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingProtection(IXLDrawingStyle style) => this._style = style;

    public bool Locked { get; set; }

    public IXLDrawingStyle SetLocked()
    {
        this.Locked = true;
        return this._style;
    }

    public IXLDrawingStyle SetLocked(bool value)
    {
        this.Locked = value;
        return this._style;
    }

    public bool LockText { get; set; }

    public IXLDrawingStyle SetLockText()
    {
        this.LockText = true;
        return this._style;
    }

    public IXLDrawingStyle SetLockText(bool value)
    {
        this.LockText = value;
        return this._style;
    }
}

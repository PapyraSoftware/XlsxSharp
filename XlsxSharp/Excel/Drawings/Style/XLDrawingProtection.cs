#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingProtection : IXLDrawingProtection
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingProtection(IXLDrawingStyle style) => this._style = style;

    public Boolean Locked { get; set; }

    public IXLDrawingStyle SetLocked()
    {
        this.Locked = true;
        return this._style;
    }

    public IXLDrawingStyle SetLocked(Boolean value)
    {
        this.Locked = value;
        return this._style;
    }

    public Boolean LockText { get; set; }

    public IXLDrawingStyle SetLockText()
    {
        this.LockText = true;
        return this._style;
    }

    public IXLDrawingStyle SetLockText(Boolean value)
    {
        this.LockText = value;
        return this._style;
    }
}

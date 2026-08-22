#nullable disable

using System;

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingProtection
{
    bool Locked { get; set; }
    bool LockText { get; set; }

    IXLDrawingStyle SetLocked();
    IXLDrawingStyle SetLocked(bool value);
    IXLDrawingStyle SetLockText();
    IXLDrawingStyle SetLockText(bool value);
}

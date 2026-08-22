#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingProtection
{
    public bool Locked { get; set; }
    public bool LockText { get; set; }

    public IXLDrawingStyle SetLocked();
    public IXLDrawingStyle SetLocked(bool value);
    public IXLDrawingStyle SetLockText();
    public IXLDrawingStyle SetLockText(bool value);
}

#nullable disable

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingFont : IXLFontBase
{
    public IXLDrawingStyle SetBold();
    public IXLDrawingStyle SetBold(bool value);
    public IXLDrawingStyle SetItalic();
    public IXLDrawingStyle SetItalic(bool value);
    public IXLDrawingStyle SetUnderline();
    public IXLDrawingStyle SetUnderline(XLFontUnderlineValues value);
    public IXLDrawingStyle SetStrikethrough();
    public IXLDrawingStyle SetStrikethrough(bool value);
    public IXLDrawingStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
    public IXLDrawingStyle SetShadow();
    public IXLDrawingStyle SetShadow(bool value);
    public IXLDrawingStyle SetFontSize(double value);
    public IXLDrawingStyle SetFontColor(XLColor value);
    public IXLDrawingStyle SetFontName(string value);
    public IXLDrawingStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value);
    public IXLDrawingStyle SetFontCharSet(XLFontCharSet value);
    public IXLDrawingStyle SetFontScheme(XLFontScheme value);
}

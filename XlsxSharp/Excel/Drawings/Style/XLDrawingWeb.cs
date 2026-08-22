using System;

namespace XlsxSharp.Excel.Drawings.Style;

internal class XLDrawingWeb : IXLDrawingWeb
{
    private readonly IXLDrawingStyle _style;

    public XLDrawingWeb(IXLDrawingStyle style)
    {
        this._style = style;
    }

    public String? AlternateText { get; set; }

    public IXLDrawingStyle SetAlternateText(String? value)
    {
        this.AlternateText = value;
        return this._style;
    }
}

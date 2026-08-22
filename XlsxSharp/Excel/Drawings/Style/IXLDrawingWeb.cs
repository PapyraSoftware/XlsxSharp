using System;

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingWeb
{
    String? AlternateText { get; set; }
    IXLDrawingStyle SetAlternateText(String? value);
}

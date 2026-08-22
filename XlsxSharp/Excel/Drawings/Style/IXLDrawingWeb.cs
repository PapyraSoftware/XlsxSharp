using System;

namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingWeb
{
    string? AlternateText { get; set; }
    IXLDrawingStyle SetAlternateText(string? value);
}

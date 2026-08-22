namespace XlsxSharp.Excel.Drawings.Style;

public interface IXLDrawingWeb
{
    public string? AlternateText { get; set; }
    public IXLDrawingStyle SetAlternateText(string? value);
}

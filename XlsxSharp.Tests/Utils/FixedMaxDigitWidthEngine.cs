using XlsxSharp.Excel;
using XlsxSharp.Excel.Drawings;
using XlsxSharp.Graphics;

namespace XlsxSharp.Tests.Utils;

/// <summary>
/// A graphic engine that reports a fixed maximum digit width and leaves everything else to the
/// default engine. Every column width of a workbook is derived from the maximum digit width of a
/// font, so a test of the conversion needs the number to be the same everywhere. Fonts of a workbook
/// are not installed on every machine and the default engine then measures whatever font it falls
/// back to, which differs from OS to OS.
/// </summary>
/// <param name="maxDigitWidth">Maximum digit width in pixels that the engine reports for every font.</param>
internal sealed class FixedMaxDigitWidthEngine(double maxDigitWidth) : IXLGraphicEngine
{
    private readonly IXLGraphicEngine _engine = DefaultGraphicEngine.Instance.Value;

    public double GetMaxDigitWidth(IXLFontBase font, double dpiX) => maxDigitWidth;

    public XLPictureInfo GetPictureInfo(Stream imageStream, XLPictureFormat expectedFormat) =>
        this._engine.GetPictureInfo(imageStream, expectedFormat);

    public double GetTextHeight(IXLFontBase font, double dpiY) =>
        this._engine.GetTextHeight(font, dpiY);

    public double GetTextWidth(string text, IXLFontBase font, double dpiX) =>
        this._engine.GetTextWidth(text, font, dpiX);

    public double GetDescent(IXLFontBase font, double dpiY) => this._engine.GetDescent(font, dpiY);

    public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi) =>
        this._engine.GetGlyphBox(graphemeCluster, font, dpi);
}

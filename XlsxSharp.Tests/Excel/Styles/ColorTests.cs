using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxSharp.Excel;
using XlsxSharp.Utils;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

public class ColorTests
{
    [Test]
    public void ColorEqualOperatorInPlace() => ClassicAssert.IsTrue(XLColor.Black == XLColor.Black);

    [Test]
    public void ColorNotEqualOperatorInPlace() =>
        ClassicAssert.IsFalse(XLColor.Black != XLColor.Black);

    [Test]
    public void ColorNamedVsHtml() =>
        ClassicAssert.IsTrue(XLColor.Black == XLColor.FromHtml("#000000"));

    [Test]
    public void DefaultStyleColorIsAutomatic()
    {
        using XLWorkbook wb = new();
        IXLWorksheet ws = wb.AddWorksheet();
        ClassicAssert.AreEqual(XLColor.Automatic, ws.FirstCell().Style.Fill.BackgroundColor);
    }

    [Test]
    public void AutomaticColorCantBeResolvedToColor()
    {
        InvalidOperationException ex = ClassicAssert.Throws<InvalidOperationException>(() =>
            _ = XLColor.Automatic.Color
        );
        ClassicAssert.AreEqual("Cannot convert automatic color to Color.", ex.Message);
    }

    [Test]
    public void CanConvertXlColorToColorType()
    {
        XLColor xlColor1 = XLColor.Red;
        XLColor xlColor2 = XLColor.FromIndex(20);
        XLColor xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
        XLColor xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

        ForegroundColor color1 = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(
            xlColor1
        );
        ForegroundColor color2 = new ForegroundColor().FromClosedXMLColor<ForegroundColor>(
            xlColor2
        );
        BackgroundColor color3 = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(
            xlColor3
        );
        BackgroundColor color4 = new BackgroundColor().FromClosedXMLColor<BackgroundColor>(
            xlColor4
        );

        ClassicAssert.AreEqual("FFFF0000", color1.Rgb.Value);
        ClassicAssert.IsNull(color1.Indexed);
        ClassicAssert.IsNull(color1.Theme);
        ClassicAssert.IsNull(color1.Tint);

        ClassicAssert.IsNull(color2.Rgb);
        ClassicAssert.AreEqual(20, color2.Indexed.Value);
        ClassicAssert.IsNull(color2.Theme);
        ClassicAssert.IsNull(color2.Tint);

        ClassicAssert.IsNull(color3.Rgb);
        ClassicAssert.IsNull(color3.Indexed);
        ClassicAssert.AreEqual(4, color3.Theme.Value);
        ClassicAssert.IsNull(color3.Tint);

        ClassicAssert.IsNull(color4.Rgb);
        ClassicAssert.IsNull(color4.Indexed);
        ClassicAssert.AreEqual(5, color4.Theme.Value);
        ClassicAssert.AreEqual(0.4, color4.Tint.Value);
    }

    [Test]
    public void CanConvertXlColorToX14ColorType()
    {
        XLColor xlColor1 = XLColor.Red;
        XLColor xlColor2 = XLColor.FromIndex(20);
        XLColor xlColor3 = XLColor.FromTheme(XLThemeColor.Accent1);
        XLColor xlColor4 = XLColor.FromTheme(XLThemeColor.Accent2, 0.4);

        X14.AxisColor color1 = new X14.AxisColor().FromClosedXMLColor<X14.AxisColor>(xlColor1);
        X14.BorderColor color2 = new X14.BorderColor().FromClosedXMLColor<X14.BorderColor>(
            xlColor2
        );
        X14.FillColor color3 = new X14.FillColor().FromClosedXMLColor<X14.FillColor>(xlColor3);
        X14.HighMarkerColor color4 =
            new X14.HighMarkerColor().FromClosedXMLColor<X14.HighMarkerColor>(xlColor4);

        ClassicAssert.AreEqual("FFFF0000", color1.Rgb.Value);
        ClassicAssert.IsNull(color1.Indexed);
        ClassicAssert.IsNull(color1.Theme);
        ClassicAssert.IsNull(color1.Tint);

        ClassicAssert.IsNull(color2.Rgb);
        ClassicAssert.AreEqual(20, color2.Indexed.Value);
        ClassicAssert.IsNull(color2.Theme);
        ClassicAssert.IsNull(color2.Tint);

        ClassicAssert.IsNull(color3.Rgb);
        ClassicAssert.IsNull(color3.Indexed);
        ClassicAssert.AreEqual(4, color3.Theme.Value);
        ClassicAssert.IsNull(color3.Tint);

        ClassicAssert.IsNull(color4.Rgb);
        ClassicAssert.IsNull(color4.Indexed);
        ClassicAssert.AreEqual(5, color4.Theme.Value);
        ClassicAssert.AreEqual(0.4, color4.Tint.Value);
    }

    [Test]
    public void CanConvertColorTypeToXlColor()
    {
        ForegroundColor color1 = new()
        {
            Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000"),
        };
        ForegroundColor color2 = new()
        {
            Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20),
        };
        BackgroundColor color3 = new() { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
        BackgroundColor color4 = new()
        {
            Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
            Tint = new DocumentFormat.OpenXml.DoubleValue(0.4),
        };

        XLColor xlColor1 = color1.ToClosedXMLColor();
        XLColor xlColor2 = color2.ToClosedXMLColor();
        XLColor xlColor3 = color3.ToClosedXMLColor();
        XLColor xlColor4 = color4.ToClosedXMLColor();

        ClassicAssert.AreEqual(XLColorType.Color, xlColor1.ColorType);
        ClassicAssert.AreEqual(XLColor.Red.Color, xlColor1.Color);

        ClassicAssert.AreEqual(XLColorType.Indexed, xlColor2.ColorType);
        ClassicAssert.AreEqual(20, xlColor2.Indexed);

        ClassicAssert.AreEqual(XLColorType.Theme, xlColor3.ColorType);
        ClassicAssert.AreEqual(XLThemeColor.Accent1, xlColor3.ThemeColor);
        ClassicAssert.AreEqual(0, xlColor3.ThemeTint, XLHelper.Epsilon);

        ClassicAssert.AreEqual(XLColorType.Theme, xlColor4.ColorType);
        ClassicAssert.AreEqual(XLThemeColor.Accent1, xlColor4.ThemeColor);
        ClassicAssert.AreEqual(0.4, xlColor4.ThemeTint, XLHelper.Epsilon);
    }

    [Test]
    public void CanConvertX14ColorTypeToXlColor()
    {
        X14.AxisColor color1 = new()
        {
            Rgb = new DocumentFormat.OpenXml.HexBinaryValue("FFFF0000"),
        };
        X14.BorderColor color2 = new()
        {
            Indexed = new DocumentFormat.OpenXml.UInt32Value((uint)20),
        };
        X14.FillColor color3 = new() { Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4) };
        X14.HighMarkerColor color4 = new()
        {
            Theme = new DocumentFormat.OpenXml.UInt32Value((uint)4),
            Tint = new DocumentFormat.OpenXml.DoubleValue(0.4),
        };

        XLColor xlColor1 = color1.ToClosedXMLColor();
        XLColor xlColor2 = color2.ToClosedXMLColor();
        XLColor xlColor3 = color3.ToClosedXMLColor();
        XLColor xlColor4 = color4.ToClosedXMLColor();

        ClassicAssert.AreEqual(XLColorType.Color, xlColor1.ColorType);
        ClassicAssert.AreEqual(XLColor.Red.Color, xlColor1.Color);

        ClassicAssert.AreEqual(XLColorType.Indexed, xlColor2.ColorType);
        ClassicAssert.AreEqual(20, xlColor2.Indexed);

        ClassicAssert.AreEqual(XLColorType.Theme, xlColor3.ColorType);
        ClassicAssert.AreEqual(XLThemeColor.Accent1, xlColor3.ThemeColor);
        ClassicAssert.AreEqual(0, xlColor3.ThemeTint, XLHelper.Epsilon);

        ClassicAssert.AreEqual(XLColorType.Theme, xlColor4.ColorType);
        ClassicAssert.AreEqual(XLThemeColor.Accent1, xlColor4.ThemeColor);
        ClassicAssert.AreEqual(0.4, xlColor4.ThemeTint, XLHelper.Epsilon);
    }

    [Test]
    public void CanParseColorWithHashAsCultureLineSeparator()
    {
        // https://github.com/XlsxSharp/XlsxSharp/issues/675
        CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");
        culture.TextInfo.ListSeparator = "#";
        Thread.CurrentThread.CurrentCulture = culture;
        XLColor color = XLColor.FromHtml("#FF008000");
        ClassicAssert.AreEqual(XLColor.Green, color);
    }

    [Test]
    [MethodDataSource(nameof(ToStringTestCases))]
    public void ToStringWorksForAllColorTypes(XLColor colorType, string expectedString) =>
        ClassicAssert.AreEqual(expectedString, colorType.ToString());

    internal static IEnumerable<(XLColor, string)> ToStringTestCases()
    {
        yield return (XLColor.FromArgb(0xFF804010), "FF804010");
        yield return (
            XLColor.FromTheme(XLThemeColor.Text1, 0.25),
            "Color Theme: Text1, Tint: 0.25"
        );
        yield return (XLColor.FromIndex(14), "Color Index: 14");
        yield return (XLColor.Automatic, "Automatic");
    }
}

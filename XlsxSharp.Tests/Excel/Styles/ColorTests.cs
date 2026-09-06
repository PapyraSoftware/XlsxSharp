using System.Globalization;
using XlsxSharp.Excel;
using XlsxSharp.Utils;

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

    public static IEnumerable<(XLColor, string)> ToStringTestCases()
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

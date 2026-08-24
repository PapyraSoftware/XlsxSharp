using System.Drawing;
using XlsxSharp.Excel;

namespace XlsxSharp.Tests.Excel.Styles;

public class XlColorTests
{
    public static IEnumerable<object[]> VmlColors
    {
        get
        {
            // Hexadecimal color
            yield return ["#F0E0D0", Color.FromArgb(0xF0, 0xE0, 0xD0)];

            // Named color
            yield return ["red", Color.Red];

            // Palette color
            yield return ["Menu [30]", Color.FromArgb(0xF0, 0xF0, 0xF0)];
            yield return ["Menu", Color.FromArgb(0xF0, 0xF0, 0xF0)];

            // Unknown/malformed color
            yield return ["#NFOBACKGROUND", Color.FromName("#NFOBACKGROUND")];
        }
    }

    [Test]
    [MethodDataSource(nameof(VmlColors))]
    public void FromVmlColorConvertsHexadecimalColors(string colorText, Color expectedColor)
    {
        XLColor color = XLColor.FromVmlColor(colorText);

        ClassicAssert.AreEqual(XLColor.FromColor(expectedColor), color);
    }
}

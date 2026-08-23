#nullable disable

using System.Drawing;
using System.Globalization;

namespace XlsxSharp.Extensions;

internal static class ColorExtensions
{
    /// <summary>
    /// Converts a <see cref="Color"/> to a hexadecimal string representation in the format "AARRGGBB".
    /// </summary>
    /// <param name="color"></param>
    /// <returns></returns>
    public static string ToHex(this Color color)
    {
        uint hexOrder = (uint)((color.B << 0) | (color.G << 8) | (color.R << 16) | (color.A << 24));
        return hexOrder.ToString("X8", CultureInfo.InvariantCulture);
    }
}

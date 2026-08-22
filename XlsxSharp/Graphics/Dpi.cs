#nullable disable

namespace XlsxSharp.Graphics;

/// <summary>
/// A DPI resolution.
/// </summary>
public readonly struct Dpi
{
    /// <summary>
    /// Horizontal DPI resolution.
    /// </summary>
    public double X { get; }

    /// <summary>
    /// Vertical DPI resolution.
    /// </summary>
    public double Y { get; }

    public Dpi(double dpiX, double dpiY)
    {
        this.X = dpiX;
        this.Y = dpiY;
    }
}

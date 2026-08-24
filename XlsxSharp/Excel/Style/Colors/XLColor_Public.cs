using System.Drawing;
using System.Globalization;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

public enum XLColorType
{
    /// <summary>
    /// Automatic color. The actual color is determined by the application depending on a context where it is used.
    /// Generally speaking, the value is resolved either as a black (e.g. border or font color) or as a white (cell
    /// or chart fill). The <see cref="XLColor.Color"/> of automatic color has no bearing on actual resolved color
    /// and should be ignored.
    /// </summary>
    Automatic,

    /// <summary>
    /// A RGB color. It can technically specify alpha component, but Excel just ignores that and marks everything
    /// as fully opaque. The color value is stored directly in <see cref="XLColor.Color"/>.
    /// </summary>
    Color,

    /// <summary>
    /// A theme color. The color value depends on a theme of a workbook and can be resolved through <see cref="IXLTheme.ResolveThemeColor"/>.
    /// </summary>
    Theme,

    /// <summary>
    /// An indexed color. Only for legacy usage, used in times when palette was common. The only semi-valid usage
    /// is for system foreground color (index 64) and system background color (index 65). The default indexed colors can
    /// be found in <see cref="XLColor.Indexed"/> and the <see cref="XLColor.Color"/> will return a value that
    /// corresponds to the default indexed color.
    /// </summary>
    Indexed,
}

public enum XLThemeColor
{
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
}

public partial class XLColor : IEquatable<XLColor>
{
    internal bool IsAuto => !this.HasValue;

    public bool HasValue { get; }

    public XLColorType ColorType => this.Key.ColorType;

    public Color Color
    {
        get
        {
            if (this.ColorType == XLColorType.Color)
            {
                return this.Key.Color;
            }

            if (this.ColorType == XLColorType.Indexed)
            {
                return IndexedColors[this.Indexed].Color;
            }

            throw new InvalidOperationException(
                $"Cannot convert {this.LcColorType} color to Color."
            );
        }
    }

    public int Indexed
    {
        get
        {
            if (this.ColorType == XLColorType.Indexed)
            {
                return this.Key.Indexed;
            }

            throw new InvalidOperationException(
                $"Cannot convert {this.LcColorType} color to indexed color."
            );
        }
    }

    public XLThemeColor ThemeColor
    {
        get
        {
            if (this.ColorType == XLColorType.Theme)
            {
                return this.Key.ThemeColor;
            }

            throw new InvalidOperationException(
                $"Cannot convert {this.LcColorType} color to theme color."
            );
        }
    }

    public double ThemeTint
    {
        get
        {
            if (this.ColorType == XLColorType.Theme)
            {
                return this.Key.ThemeTint;
            }

            if (this.ColorType == XLColorType.Indexed)
            {
                throw new InvalidOperationException(
                    "Cannot extract theme tint from an indexed color."
                );
            }

            return this.Color.A / 255.0;
        }
    }

    #region IEquatable<XLColor> Members

    public bool Equals(XLColor other) => this.Key == other.Key;

    #endregion IEquatable<XLColor> Members

    public override bool Equals(object obj) => this.Equals((XLColor)obj);

    public override int GetHashCode()
    {
        int hashCode = 229333804;
        hashCode = hashCode * -1521134295 + this.HasValue.GetHashCode();
        hashCode = hashCode * -1521134295 + this.Key.GetHashCode();
        return hashCode;
    }

    public override string ToString()
    {
        if (this.ColorType == XLColorType.Color)
        {
            return this.Color.ToHex();
        }

        if (this.ColorType == XLColorType.Theme)
        {
            return $"Color Theme: {this.ThemeColor}, Tint: {this.ThemeTint.ToString(CultureInfo.InvariantCulture)}";
        }

        if (this.ColorType == XLColorType.Automatic)
        {
            return "Automatic";
        }

        return "Color Index: " + this.Indexed;
    }

    public static bool operator ==(XLColor? left, XLColor? right)
    {
        // If both are null, or both are same instance, return true.
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        // If one is null, but not both, return false.
        if ((left as object) == null || (right as object) == null)
        {
            return false;
        }

        return left.Equals(right);
    }

    public static bool operator !=(XLColor? left, XLColor? right) => !(left == right);
}

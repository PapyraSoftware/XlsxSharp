#nullable disable

namespace XlsxSharp.Excel;

internal struct XLColorKey : IEquatable<XLColorKey>
{
    public XLColorType ColorType { get; set; }

    public System.Drawing.Color Color { get; set; }

    public int Indexed { get; set; }

    public XLThemeColor ThemeColor { get; set; }

    public double ThemeTint { get; set; }

    public override int GetHashCode()
    {
        int hashCode = -331517974;
        hashCode = hashCode * -1521134295 + (int)this.ColorType;
        hashCode =
            hashCode * -1521134295 + (this.ColorType == XLColorType.Indexed ? this.Indexed : 0);
        hashCode =
            hashCode * -1521134295
            + (this.ColorType == XLColorType.Theme ? (int)this.ThemeColor : 0);
        hashCode =
            hashCode * -1521134295
            + (this.ColorType == XLColorType.Theme ? this.ThemeTint.GetHashCode() : 0);
        hashCode =
            hashCode * -1521134295
            + (this.ColorType == XLColorType.Color ? this.Color.ToArgb() : 0);
        return hashCode;
    }

    public bool Equals(XLColorKey other)
    {
        if (this.ColorType == other.ColorType)
        {
            if (this.ColorType == XLColorType.Color)
            {
                // .NET Color.Equals() will return false for Color.FromArgb(255, 255, 255, 255) == Color.White
                // Therefore we compare the ToArgb() values
                return this.Color.ToArgb() == other.Color.ToArgb();
            }
            if (this.ColorType == XLColorType.Theme)
            {
                return this.ThemeColor == other.ThemeColor
                    && Math.Abs(this.ThemeTint - other.ThemeTint) < XlsxSharp.XLHelper.Epsilon;
            }
            return this.Indexed == other.Indexed;
        }

        return false;
    }

    public override bool Equals(object obj)
    {
        if (obj is XLColorKey)
        {
            return this.Equals((XLColorKey)obj);
        }

        return base.Equals(obj);
    }

    public override string ToString()
    {
        switch (this.ColorType)
        {
            case XLColorType.Color:
                return this.Color.ToString();
            case XLColorType.Theme:
                return $"{this.ThemeColor} ({this.ThemeTint})";
            case XLColorType.Indexed:
                return $"Indexed: {this.Indexed}";
            default:
                return base.ToString();
        }
    }

    public static bool operator ==(XLColorKey left, XLColorKey right) => left.Equals(right);

    public static bool operator !=(XLColorKey left, XLColorKey right) => !(left.Equals(right));
}

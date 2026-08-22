#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace XlsxSharp.Excel;

internal class XLSparklineStyle : IXLSparklineStyle, IEquatable<XLSparklineStyle>
{
    #region Public Properties

    public XLColor FirstMarkerColor { get; set; }

    public XLColor HighMarkerColor { get; set; }

    public XLColor LastMarkerColor { get; set; }

    public XLColor LowMarkerColor { get; set; }

    public XLColor MarkersColor { get; set; }

    public XLColor NegativeColor { get; set; }

    public XLColor SeriesColor { get; set; }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineStyle SetFirstMarkerColor(XLColor value)
    {
        this.FirstMarkerColor = value;
        return this;
    }

    public IXLSparklineStyle SetHighMarkerColor(XLColor value)
    {
        this.HighMarkerColor = value;
        return this;
    }

    public IXLSparklineStyle SetLastMarkerColor(XLColor value)
    {
        this.LastMarkerColor = value;
        return this;
    }

    public IXLSparklineStyle SetLowMarkerColor(XLColor value)
    {
        this.LowMarkerColor = value;
        return this;
    }

    public IXLSparklineStyle SetMarkersColor(XLColor value)
    {
        this.MarkersColor = value;
        return this;
    }

    public IXLSparklineStyle SetNegativeColor(XLColor value)
    {
        this.NegativeColor = value;
        return this;
    }

    public IXLSparklineStyle SetSeriesColor(XLColor value)
    {
        this.SeriesColor = value;
        return this;
    }

    #endregion Public Methods

    public static void Copy(IXLSparklineStyle from, IXLSparklineStyle to)
    {
        to.FirstMarkerColor = from.FirstMarkerColor;
        to.HighMarkerColor = from.HighMarkerColor;
        to.LastMarkerColor = from.LastMarkerColor;
        to.LowMarkerColor = from.LowMarkerColor;
        to.MarkersColor = from.MarkersColor;
        to.NegativeColor = from.NegativeColor;
        to.SeriesColor = from.SeriesColor;
    }

    #region IEquatable implementation

    /// <summary>Returns a value that indicates whether two <see cref="T:XLSparklineStyle" /> objects have different values.</summary>
    /// <param name="left">The first value to compare.</param>
    /// <param name="right">The second value to compare.</param>
    /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
    public static bool operator !=(XLSparklineStyle left, XLSparklineStyle right) =>
        !Equals(left, right);

    /// <summary>Returns a value that indicates whether the values of two <see cref="T:XLSparklineStyle" /> objects are equal.</summary>
    /// <param name="left">The first value to compare.</param>
    /// <param name="right">The second value to compare.</param>
    /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
    public static bool operator ==(XLSparklineStyle left, XLSparklineStyle right) =>
        Equals(left, right);

    /// <summary>Indicates whether the current object is equal to another object of the same type.</summary>
    /// <param name="other">An object to compare with this object.</param>
    /// <returns>true if the current object is equal to the <paramref name="other">other</paramref> parameter; otherwise, false.</returns>
    public bool Equals(XLSparklineStyle other)
    {
        if (ReferenceEquals(null, other))
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return this.FirstMarkerColor.Equals(other.FirstMarkerColor)
            && this.HighMarkerColor.Equals(other.HighMarkerColor)
            && this.LastMarkerColor.Equals(other.LastMarkerColor)
            && this.LowMarkerColor.Equals(other.LowMarkerColor)
            && this.MarkersColor.Equals(other.MarkersColor)
            && this.NegativeColor.Equals(other.NegativeColor)
            && this.SeriesColor.Equals(other.SeriesColor);
    }

    /// <summary>Determines whether the specified object is equal to the current object.</summary>
    /// <param name="obj">The object to compare with the current object.</param>
    /// <returns>true if the specified object  is equal to the current object; otherwise, false.</returns>
    public override bool Equals(object obj)
    {
        if (ReferenceEquals(null, obj))
        {
            return false;
        }

        if (ReferenceEquals(this, obj))
        {
            return true;
        }

        if (obj.GetType() != typeof(XLSparklineStyle))
        {
            return false;
        }

        return this.Equals((XLSparklineStyle)obj);
    }

    /// <summary>Serves as the default hash function.</summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode() =>
        HashCode.Combine(
            this.FirstMarkerColor,
            this.HighMarkerColor,
            this.LastMarkerColor,
            this.LowMarkerColor,
            this.MarkersColor,
            this.NegativeColor,
            this.SeriesColor
        );

    #endregion IEquatable implementation
}

namespace XlsxSharp.Excel.Formatting;

/// <summary>
/// Border line properties.
/// </summary>
internal readonly record struct XLBorderLine(XLColor Color, XLBorderStyleValues Style)
{
    /// <summary>
    /// Default XML attribute values.
    /// </summary>
    internal static readonly XLBorderLine None = new(XLColor.Automatic, XLBorderStyleValues.None);

    /// <summary>
    /// Is the border line visible?
    /// </summary>
    internal bool IsVisible => this.Style != XLBorderStyleValues.None;
}

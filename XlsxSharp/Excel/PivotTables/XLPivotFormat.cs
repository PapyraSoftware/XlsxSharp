using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel;

/// <summary>
/// A description of formatting that should be applied to a <see cref="XLPivotTable"/>.
/// </summary>
internal class XLPivotFormat : IXLDxfContainer
{
    internal XLPivotFormat(XLPivotArea pivotArea)
    {
        this.PivotArea = pivotArea;
    }

    /// <summary>
    /// Pivot area that should be formatted.
    /// </summary>
    internal XLPivotArea PivotArea { get; }

    /// <summary>
    /// Should the formatting (determined by <see cref="FormatValue"/>) be applied or not?
    /// </summary>
    internal XLPivotFormatAction Action { get; init; } = XLPivotFormatAction.Formatting;

    /// <summary>
    /// Differential formatting to apply to the <see cref="PivotArea"/>. It can be empty, e.g. if
    /// <see cref="Action"/> is blank.
    /// </summary>
    public XLDxfValue? FormatValue { get; set; }
}

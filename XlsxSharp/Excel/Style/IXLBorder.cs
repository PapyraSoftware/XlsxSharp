#nullable disable

using XlsxSharp.Excel.Rows;

namespace XlsxSharp.Excel;

public enum XLBorderStyleValues
{
    DashDot,
    DashDotDot,
    Dashed,
    Dotted,
    Double,
    Hair,
    Medium,
    MediumDashDot,
    MediumDashDotDot,
    MediumDashed,
    None,
    SlantDashDot,
    Thick,
    Thin,
}

/// <summary>
/// <para>
/// The interface is used across many different objects. The value returned by properties is
/// defined only for <see cref="IXLCell"/>, <see cref="IXLRow"/>, <see cref="IXLColumn"/>,
/// <see cref="IXLWorksheet"/> and <see cref="XLWorkbook"/>. The returned value is undefined
/// for other <see cref="IXLRangeBase"/> objects that can contain multiple different property
/// values (e.g. <see cref="IXLRange"/> can contain multiple cells, <see cref="IXLColumns"/>
/// can contains multiple columns).
/// </para>
/// </summary>
public interface IXLBorder : IEquatable<IXLBorder>
{
    public XLBorderStyleValues OutsideBorder { set; }

    public XLColor OutsideBorderColor { set; }

    public XLBorderStyleValues InsideBorder { set; }

    public XLColor InsideBorderColor { set; }

    /// <summary>
    /// Get or set style of the left border.
    /// </summary>
    /// <remarks>
    /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="LeftBorderColor"/>
    /// is set to the default border color.
    /// </remarks>
    public XLBorderStyleValues LeftBorder { get; set; }

    /// <summary>
    /// Get or set color of the left border.
    /// </summary>
    /// <remarks>
    /// The color can be set only when the border is visible (=<see cref="LeftBorder"/>
    /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
    /// </remarks>
    public XLColor LeftBorderColor { get; set; }

    /// <summary>
    /// Get or set style of the right border.
    /// </summary>
    /// <remarks>
    /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="RightBorderColor"/>
    /// is set to the default border color.
    /// </remarks>
    public XLBorderStyleValues RightBorder { get; set; }

    /// <summary>
    /// Get or set color of the right border.
    /// </summary>
    /// <remarks>
    /// The color can be set only when the border is visible (=<see cref="RightBorder"/>
    /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
    /// </remarks>
    public XLColor RightBorderColor { get; set; }

    /// <summary>
    /// Get or set style of the top border.
    /// </summary>
    /// <remarks>
    /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="TopBorderColor"/>
    /// is set to the default border color.
    /// </remarks>
    public XLBorderStyleValues TopBorder { get; set; }

    /// <summary>
    /// Get or set color of the top border.
    /// </summary>
    /// <remarks>
    /// The color can be set only when the border is visible (=<see cref="TopBorder"/>
    /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
    /// </remarks>
    public XLColor TopBorderColor { get; set; }

    /// <summary>
    /// Get or set style of the bottom border.
    /// </summary>
    /// <remarks>
    /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="BottomBorderColor"/>
    /// is set to the default border color.
    /// </remarks>
    public XLBorderStyleValues BottomBorder { get; set; }

    /// <summary>
    /// Get or set color of the bottom border.
    /// </summary>
    /// <remarks>
    /// The color can be set only when the border is visible (=<see cref="BottomBorder"/>
    /// is not <see cref="XLBorderStyleValues.None"/>). Set style first, then color.
    /// </remarks>
    public XLColor BottomBorderColor { get; set; }

    public bool DiagonalUp { get; set; }

    public bool DiagonalDown { get; set; }

    /// <summary>
    /// Get or set style of the diagonal border.
    /// </summary>
    /// <remarks>
    /// When style is set to <see cref="XLBorderStyleValues.None"/>, the <see cref="DiagonalBorderColor"/>
    /// is set to the default border color.
    /// </remarks>
    public XLBorderStyleValues DiagonalBorder { get; set; }

    /// <summary>
    /// Get or set color of the diagonal border.
    /// </summary>
    /// <remarks>
    /// The color can be set only when the border line is can be visible
    /// (=<see cref="DiagonalBorder"/> is not <see cref="XLBorderStyleValues.None"/>).
    /// Set style first, then color.
    /// </remarks>
    public XLColor DiagonalBorderColor { get; set; }

    public IXLStyle SetOutsideBorder(XLBorderStyleValues value);

    public IXLStyle SetOutsideBorderColor(XLColor value);

    public IXLStyle SetInsideBorder(XLBorderStyleValues value);

    public IXLStyle SetInsideBorderColor(XLColor value);

    /// <summary>
    /// Set style of the left border.
    /// </summary>
    /// <inheritdoc cref="LeftBorder"/>
    public IXLStyle SetLeftBorder(XLBorderStyleValues value);

    /// <summary>
    /// Set color of the left border.
    /// </summary>
    /// <inheritdoc cref="LeftBorderColor"/>
    public IXLStyle SetLeftBorderColor(XLColor value);

    /// <summary>
    /// Set style of the right border.
    /// </summary>
    /// <inheritdoc cref="RightBorder"/>
    public IXLStyle SetRightBorder(XLBorderStyleValues value);

    /// <summary>
    /// Set color of the right border.
    /// </summary>
    /// <inheritdoc cref="RightBorderColor"/>
    public IXLStyle SetRightBorderColor(XLColor value);

    /// <summary>
    /// Set style of the top border.
    /// </summary>
    /// <inheritdoc cref="TopBorder"/>
    public IXLStyle SetTopBorder(XLBorderStyleValues value);

    /// <summary>
    /// Set color of the top border.
    /// </summary>
    /// <inheritdoc cref="TopBorderColor"/>
    public IXLStyle SetTopBorderColor(XLColor value);

    /// <summary>
    /// Set style of the bottom border.
    /// </summary>
    /// <inheritdoc cref="BottomBorder"/>
    public IXLStyle SetBottomBorder(XLBorderStyleValues value);

    /// <summary>
    /// Set color of the bottom border.
    /// </summary>
    /// <inheritdoc cref="BottomBorderColor"/>
    public IXLStyle SetBottomBorderColor(XLColor value);

    public IXLStyle SetDiagonalUp();
    public IXLStyle SetDiagonalUp(bool value);

    public IXLStyle SetDiagonalDown();
    public IXLStyle SetDiagonalDown(bool value);

    /// <summary>
    /// Set style of the diagonal border.
    /// </summary>
    /// <inheritdoc cref="DiagonalBorder"/>
    public IXLStyle SetDiagonalBorder(XLBorderStyleValues value);

    /// <summary>
    /// Set color of the diagonal border.
    /// </summary>
    /// <inheritdoc cref="DiagonalBorderColor"/>
    public IXLStyle SetDiagonalBorderColor(XLColor value);
}

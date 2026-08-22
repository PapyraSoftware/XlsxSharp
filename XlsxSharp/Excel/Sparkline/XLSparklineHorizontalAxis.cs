#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;

namespace XlsxSharp.Excel;

internal class XLSparklineHorizontalAxis : IXLSparklineHorizontalAxis
{
    #region Public Properties

    public XLColor Color { get; set; }

    public bool DateAxis => this.SparklineGroup.DateRange != null;

    public bool IsVisible { get; set; }

    public bool RightToLeft { get; set; }

    public IXLSparklineGroup SparklineGroup { get; }

    #endregion Public Properties

    #region Public Constructors

    public XLSparklineHorizontalAxis(IXLSparklineGroup sparklineGroup)
    {
        this.SparklineGroup =
            sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));
    }

    #endregion Public Constructors

    #region Public Methods

    public IXLSparklineHorizontalAxis SetColor(XLColor value)
    {
        this.Color = value ?? throw new ArgumentNullException(nameof(value));
        return this;
    }

    public IXLSparklineHorizontalAxis SetRightToLeft(bool value)
    {
        this.RightToLeft = value;
        return this;
    }

    public IXLSparklineHorizontalAxis SetVisible(bool value)
    {
        this.IsVisible = value;
        return this;
    }

    #endregion Public Methods

    public static void Copy(IXLSparklineHorizontalAxis from, IXLSparklineHorizontalAxis to)
    {
        to.Color = from.Color;
        to.IsVisible = from.IsVisible;
        to.RightToLeft = from.RightToLeft;
    }
}

#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace XlsxSharp.Excel;

public interface IXLSparklineHorizontalAxis
{
    #region Public Properties

    public XLColor Color { get; set; }

    public bool DateAxis { get; }

    public bool IsVisible { get; set; }

    public bool RightToLeft { get; set; }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineHorizontalAxis SetColor(XLColor value);

    public IXLSparklineHorizontalAxis SetRightToLeft(bool value);

    public IXLSparklineHorizontalAxis SetVisible(bool value);

    #endregion Public Methods
}

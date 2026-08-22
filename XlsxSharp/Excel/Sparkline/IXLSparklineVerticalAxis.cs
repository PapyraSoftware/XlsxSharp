#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace XlsxSharp.Excel;

public interface IXLSparklineVerticalAxis
{
    #region Public Properties

    public double? ManualMax { get; set; }

    public double? ManualMin { get; set; }

    public XLSparklineAxisMinMax MaxAxisType { get; set; }

    public XLSparklineAxisMinMax MinAxisType { get; set; }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineVerticalAxis SetManualMax(double? value);

    public IXLSparklineVerticalAxis SetManualMin(double? value);

    public IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax value);

    public IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax value);

    #endregion Public Methods
}

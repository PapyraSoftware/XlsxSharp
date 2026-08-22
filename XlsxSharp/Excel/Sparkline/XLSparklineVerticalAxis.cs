#nullable disable

using System;

namespace XlsxSharp.Excel;

internal class XLSparklineVerticalAxis : IXLSparklineVerticalAxis
{
    #region Public Properties

    public double? ManualMax
    {
        get => this._manualMax;
        set => this.SetManualMax(value);
    }

    public double? ManualMin
    {
        get => this._manualMin;
        set => this.SetManualMin(value);
    }

    public XLSparklineAxisMinMax MaxAxisType
    {
        get => this._maxAxisType;
        set => this.SetMaxAxisType(value);
    }

    public XLSparklineAxisMinMax MinAxisType
    {
        get => this._minAxisType;
        set => this.SetMinAxisType(value);
    }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineVerticalAxis SetManualMax(double? manualMax)
    {
        if (manualMax != null)
        {
            this.MaxAxisType = XLSparklineAxisMinMax.Custom;
        }

        this._manualMax = manualMax;
        return this;
    }

    public IXLSparklineVerticalAxis SetManualMin(double? manualMin)
    {
        if (manualMin != null)
        {
            this.MinAxisType = XLSparklineAxisMinMax.Custom;
        }

        this._manualMin = manualMin;
        return this;
    }

    public IXLSparklineVerticalAxis SetMaxAxisType(XLSparklineAxisMinMax maxAxisType)
    {
        if (maxAxisType != XLSparklineAxisMinMax.Custom)
        {
            this._manualMax = null;
        }

        this._maxAxisType = maxAxisType;
        return this;
    }

    public IXLSparklineVerticalAxis SetMinAxisType(XLSparklineAxisMinMax minAxisType)
    {
        if (minAxisType != XLSparklineAxisMinMax.Custom)
        {
            this._manualMin = null;
        }

        this._minAxisType = minAxisType;
        return this;
    }

    #endregion Public Methods

    #region Private Fields

    private double? _manualMax;
    private double? _manualMin;
    private XLSparklineAxisMinMax _maxAxisType;
    private XLSparklineAxisMinMax _minAxisType;

    #endregion Private Fields

    public IXLSparklineGroup SparklineGroup { get; }

    public XLSparklineVerticalAxis(IXLSparklineGroup sparklineGroup) =>
        this.SparklineGroup =
            sparklineGroup ?? throw new ArgumentNullException(nameof(sparklineGroup));

    public static void Copy(IXLSparklineVerticalAxis from, IXLSparklineVerticalAxis to)
    {
        to.ManualMax = from.ManualMax;
        to.ManualMin = from.ManualMin;
        to.MaxAxisType = from.MaxAxisType;
        to.MinAxisType = from.MinAxisType;
    }
}

#nullable disable

// Keep this file CodeMaid organised and cleaned

namespace XlsxSharp.Excel;

public interface IXLSparklineStyle
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

    public IXLSparklineStyle SetFirstMarkerColor(XLColor value);

    public IXLSparklineStyle SetHighMarkerColor(XLColor value);

    public IXLSparklineStyle SetLastMarkerColor(XLColor value);

    public IXLSparklineStyle SetLowMarkerColor(XLColor value);

    public IXLSparklineStyle SetMarkersColor(XLColor value);

    public IXLSparklineStyle SetNegativeColor(XLColor value);

    public IXLSparklineStyle SetSeriesColor(XLColor value);

    #endregion Public Methods
}

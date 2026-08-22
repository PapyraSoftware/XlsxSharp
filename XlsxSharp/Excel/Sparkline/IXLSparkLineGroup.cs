#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;

namespace XlsxSharp.Excel;

public enum XLDisplayBlanksAsValues
{
    Interpolate = 0,
    NotPlotted = 1,
    Zero = 2,
}

public enum XLSparklineAxisMinMax
{
    Automatic = 0,
    SameForAll = 1,
    Custom = 2,
}

[Flags]
public enum XLSparklineMarkers
{
    None = 0,
    HighPoint = 1 << 1,
    LowPoint = 1 << 2,
    FirstPoint = 1 << 3,
    LastPoint = 1 << 4,
    NegativePoints = 1 << 5,
    Markers = 1 << 6,
    All = HighPoint | LowPoint | FirstPoint | LastPoint | NegativePoints | Markers,
}

public enum XLSparklineType
{
    Line = 0,
    Column = 1,
    Stacked = 2,
}

public interface IXLSparklineGroup : IEnumerable<IXLSparkline>
{
    public IXLRange DateRange { get; set; }

    public XLDisplayBlanksAsValues DisplayEmptyCellsAs { get; set; }

    public bool DisplayHidden { get; set; }

    public IXLSparklineHorizontalAxis HorizontalAxis { get; }

    public double LineWeight { get; set; }

    public XLSparklineMarkers ShowMarkers { get; set; }

    public IXLSparklineStyle Style { get; set; }

    public XLSparklineType Type { get; set; }

    public IXLSparklineVerticalAxis VerticalAxis { get; }

    public IXLWorksheet Worksheet { get; }

    public IXLSparkline Add(IXLCell location, IXLRange sourceData);

    public IEnumerable<IXLSparkline> Add(IXLRange locationRange, IXLRange sourceDataRange);

    public IEnumerable<IXLSparkline> Add(string locationRangeAddress, string sourceDataAddress);

    public void CopyFrom(IXLSparklineGroup sparklineGroup);

    /// <summary>
    /// Copy this sparkline group to the specified worksheet
    /// </summary>
    /// <param name="targetSheet">The worksheet to copy this sparkline group to</param>
    public IXLSparklineGroup CopyTo(IXLWorksheet targetSheet);

    public IXLSparkline GetSparkline(IXLCell cell);

    public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase searchRange);

    public void Remove(IXLCell cell);

    public void Remove(IXLSparkline sparkline);

    public void RemoveAll();

    public IXLSparklineGroup SetDateRange(IXLRange value);

    public IXLSparklineGroup SetDisplayEmptyCellsAs(XLDisplayBlanksAsValues value);

    public IXLSparklineGroup SetDisplayHidden(bool value);

    public IXLSparklineGroup SetLineWeight(double value);

    public IXLSparklineGroup SetShowMarkers(XLSparklineMarkers value);

    public IXLSparklineGroup SetStyle(IXLSparklineStyle value);

    public IXLSparklineGroup SetType(XLSparklineType value);
}

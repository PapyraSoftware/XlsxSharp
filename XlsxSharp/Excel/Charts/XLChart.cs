#nullable disable

using System;
using System.Collections.Generic;
using System.Linq;
using XlsxSharp.Excel.Drawings;

namespace XlsxSharp.Excel.Charts;

internal enum XLChartTypeCategory
{
    Bar3D,
}

internal enum XLBarOrientation
{
    Vertical,
    Horizontal,
}

internal enum XLBarGrouping
{
    Clustered,
    Percent,
    Stacked,
    Standard,
}

internal class XLChart : XLDrawing<IXLChart>, IXLChart
{
    internal IXLWorksheet worksheet;

    public XLChart(XLWorksheet worksheet)
    {
        this.Container = this;
        this.worksheet = worksheet;
        int zOrder;
        if (worksheet.Charts.Any())
        {
            zOrder = worksheet.Charts.Max(c => c.ZOrder) + 1;
        }
        else
        {
            zOrder = 1;
        }

        this.ZOrder = zOrder;
        this.ShapeId = worksheet.Workbook.ShapeIdManager.GetNext();
        this.RightAngleAxes = true;
    }

    public bool RightAngleAxes { get; set; }

    public IXLChart SetRightAngleAxes()
    {
        this.RightAngleAxes = true;
        return this;
    }

    public IXLChart SetRightAngleAxes(bool rightAngleAxes)
    {
        this.RightAngleAxes = rightAngleAxes;
        return this;
    }

    public XLChartType ChartType { get; set; }

    public IXLChart SetChartType(XLChartType chartType)
    {
        this.ChartType = chartType;
        return this;
    }

    public XLChartTypeCategory ChartTypeCategory
    {
        get
        {
            if (this.Bar3DCharts.Contains(this.ChartType))
            {
                return XLChartTypeCategory.Bar3D;
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }

    private HashSet<XLChartType> Bar3DCharts =
    [
        XLChartType.BarClustered3D,
        XLChartType.BarStacked100Percent3D,
        XLChartType.BarStacked3D,
        XLChartType.Column3D,
        XLChartType.ColumnClustered3D,
        XLChartType.ColumnStacked100Percent3D,
        XLChartType.ColumnStacked3D,
    ];

    public XLBarOrientation BarOrientation
    {
        get
        {
            if (this.HorizontalCharts.Contains(this.ChartType))
            {
                return XLBarOrientation.Horizontal;
            }
            else
            {
                return XLBarOrientation.Vertical;
            }
        }
    }

    private HashSet<XLChartType> HorizontalCharts =
    [
        XLChartType.BarClustered,
        XLChartType.BarClustered3D,
        XLChartType.BarStacked,
        XLChartType.BarStacked100Percent,
        XLChartType.BarStacked100Percent3D,
        XLChartType.BarStacked3D,
        XLChartType.ConeHorizontalClustered,
        XLChartType.ConeHorizontalStacked,
        XLChartType.ConeHorizontalStacked100Percent,
        XLChartType.CylinderHorizontalClustered,
        XLChartType.CylinderHorizontalStacked,
        XLChartType.CylinderHorizontalStacked100Percent,
        XLChartType.PyramidHorizontalClustered,
        XLChartType.PyramidHorizontalStacked,
        XLChartType.PyramidHorizontalStacked100Percent,
    ];

    public XLBarGrouping BarGrouping
    {
        get
        {
            if (this.ClusteredCharts.Contains(this.ChartType))
            {
                return XLBarGrouping.Clustered;
            }
            else if (this.PercentCharts.Contains(this.ChartType))
            {
                return XLBarGrouping.Percent;
            }
            else if (this.StackedCharts.Contains(this.ChartType))
            {
                return XLBarGrouping.Stacked;
            }
            else
            {
                return XLBarGrouping.Standard;
            }
        }
    }

    public HashSet<XLChartType> ClusteredCharts =
    [
        XLChartType.BarClustered,
        XLChartType.BarClustered3D,
        XLChartType.ColumnClustered,
        XLChartType.ColumnClustered3D,
        XLChartType.ConeClustered,
        XLChartType.ConeHorizontalClustered,
        XLChartType.CylinderClustered,
        XLChartType.CylinderHorizontalClustered,
        XLChartType.PyramidClustered,
        XLChartType.PyramidHorizontalClustered,
    ];

    public HashSet<XLChartType> PercentCharts =
    [
        XLChartType.AreaStacked100Percent,
        XLChartType.AreaStacked100Percent3D,
        XLChartType.BarStacked100Percent,
        XLChartType.BarStacked100Percent3D,
        XLChartType.ColumnStacked100Percent,
        XLChartType.ColumnStacked100Percent3D,
        XLChartType.ConeHorizontalStacked100Percent,
        XLChartType.ConeStacked100Percent,
        XLChartType.CylinderHorizontalStacked100Percent,
        XLChartType.CylinderStacked100Percent,
        XLChartType.LineStacked100Percent,
        XLChartType.LineWithMarkersStacked100Percent,
        XLChartType.PyramidHorizontalStacked100Percent,
        XLChartType.PyramidStacked100Percent,
    ];

    public HashSet<XLChartType> StackedCharts =
    [
        XLChartType.AreaStacked,
        XLChartType.AreaStacked3D,
        XLChartType.BarStacked,
        XLChartType.BarStacked3D,
        XLChartType.ColumnStacked,
        XLChartType.ColumnStacked3D,
        XLChartType.ConeHorizontalStacked,
        XLChartType.ConeStacked,
        XLChartType.CylinderHorizontalStacked,
        XLChartType.CylinderStacked,
        XLChartType.LineStacked,
        XLChartType.LineWithMarkersStacked,
        XLChartType.PyramidHorizontalStacked,
        XLChartType.PyramidStacked,
    ];
}

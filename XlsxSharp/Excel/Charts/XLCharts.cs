#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.Charts;

internal class XLCharts : IXLCharts
{
    private List<IXLChart> charts = [];

    public IEnumerator<IXLChart> GetEnumerator()
    {
        return this.charts.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public void Add(IXLChart chart)
    {
        this.charts.Add(chart);
    }
}

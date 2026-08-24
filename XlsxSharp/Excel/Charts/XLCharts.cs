#nullable disable

namespace XlsxSharp.Excel.Charts;

internal class XLCharts : IXLCharts
{
    private List<IXLChart> charts = [];

    public IEnumerator<IXLChart> GetEnumerator() => this.charts.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();

    public void Add(IXLChart chart) => this.charts.Add(chart);
}

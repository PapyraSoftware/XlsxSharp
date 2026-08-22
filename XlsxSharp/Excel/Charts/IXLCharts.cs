#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.Charts;

public interface IXLCharts : IEnumerable<IXLChart>
{
    public void Add(IXLChart chart);
}

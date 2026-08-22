#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.Charts;

public interface IXLCharts : IEnumerable<IXLChart>
{
    void Add(IXLChart chart);
}

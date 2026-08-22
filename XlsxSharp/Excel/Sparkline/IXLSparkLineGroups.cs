#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel;

public interface IXLSparklineGroups : IEnumerable<IXLSparklineGroup>
{
    #region Public Properties

    public IXLWorksheet Worksheet { get; }

    #endregion Public Properties

    #region Public Methods

    public IXLSparklineGroup Add(IXLSparklineGroup sparklineGroup);
    public IXLSparklineGroup Add(string locationAddress, string sourceDataAddress);
    public IXLSparklineGroup Add(IXLCell location, IXLRange sourceData);
    public IXLSparklineGroup Add(IXLRange locationRange, IXLRange sourceDataRange);

    public void CopyTo(IXLWorksheet targetSheet);

    public IXLSparkline GetSparkline(IXLCell cell);
    public IEnumerable<IXLSparkline> GetSparklines(IXLRangeBase rangeBase);

    public void Remove(IXLCell cell);
    public void Remove(IXLRangeBase range);
    public void Remove(IXLSparklineGroup sparklineGroup);
    public void RemoveAll();

    #endregion Public Methods
}

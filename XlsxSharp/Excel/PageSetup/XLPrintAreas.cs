#nullable disable

using System.Collections.Generic;

namespace XlsxSharp.Excel.PageSetup;

internal class XLPrintAreas : IXLPrintAreas
{
    List<IXLRange> ranges = [];
    private XLWorksheet worksheet;

    public XLPrintAreas(XLWorksheet worksheet) => this.worksheet = worksheet;

    public XLPrintAreas(XLPrintAreas defaultPrintAreas, XLWorksheet worksheet)
    {
        this.ranges = [.. defaultPrintAreas.ranges];
        this.worksheet = worksheet;
    }

    public void Clear() => this.ranges.Clear();

    public void Add(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn) =>
        this.ranges.Add(
            this.worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn)
        );

    public void Add(string rangeAddress) => this.ranges.Add(this.worksheet.Range(rangeAddress));

    public void Add(string firstCellAddress, string lastCellAddress) =>
        this.ranges.Add(this.worksheet.Range(firstCellAddress, lastCellAddress));

    public void Add(IXLAddress firstCellAddress, IXLAddress lastCellAddress) =>
        this.ranges.Add(this.worksheet.Range(firstCellAddress, lastCellAddress));

    public IEnumerator<IXLRange> GetEnumerator() => this.ranges.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        this.GetEnumerator();
}

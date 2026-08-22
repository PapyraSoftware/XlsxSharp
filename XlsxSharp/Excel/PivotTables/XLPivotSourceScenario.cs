using System;

namespace XlsxSharp.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes uses scenarios in the workbook to
/// create data.
/// </summary>
internal sealed class XLPivotSourceScenario : IXLPivotSource
{
    public bool Equals(IXLPivotSource other) => other is XLPivotSourceScenario;

    public override bool Equals(object? obj) => obj is IXLPivotSource other && this.Equals(other);

    public override int GetHashCode() => 0;

    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea) =>
        throw new NotImplementedException("Scenario pivot cache data source is not supported.");
}

#nullable disable

namespace XlsxSharp.Excel;

public enum XLOutlineSummaryVLocation
{
    Top,
    Bottom,
};

public enum XLOutlineSummaryHLocation
{
    Left,
    Right,
};

public interface IXLOutline
{
    public XLOutlineSummaryVLocation SummaryVLocation { get; set; }
    public XLOutlineSummaryHLocation SummaryHLocation { get; set; }
}

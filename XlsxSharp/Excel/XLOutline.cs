#nullable disable

namespace XlsxSharp.Excel;

internal class XLOutline : IXLOutline
{
    public XLOutline(IXLOutline outline)
    {
        if (outline != null)
        {
            this.SummaryHLocation = outline.SummaryHLocation;
            this.SummaryVLocation = outline.SummaryVLocation;
        }
    }

    public XLOutlineSummaryVLocation SummaryVLocation { get; set; }
    public XLOutlineSummaryHLocation SummaryHLocation { get; set; }
}

#nullable disable

namespace XlsxSharp.Excel.PageSetup;

internal class XLMargins : IXLMargins
{
    public double Left { get; set; }
    public double Right { get; set; }
    public double Top { get; set; }
    public double Bottom { get; set; }
    public double Header { get; set; }
    public double Footer { get; set; }

    public IXLMargins SetLeft(double value)
    {
        this.Left = value;
        return this;
    }

    public IXLMargins SetRight(double value)
    {
        this.Right = value;
        return this;
    }

    public IXLMargins SetTop(double value)
    {
        this.Top = value;
        return this;
    }

    public IXLMargins SetBottom(double value)
    {
        this.Bottom = value;
        return this;
    }

    public IXLMargins SetHeader(double value)
    {
        this.Header = value;
        return this;
    }

    public IXLMargins SetFooter(double value)
    {
        this.Footer = value;
        return this;
    }
}

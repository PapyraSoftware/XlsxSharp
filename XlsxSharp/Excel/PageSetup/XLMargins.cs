#nullable disable

using System;

namespace XlsxSharp.Excel.PageSetup;

internal class XLMargins : IXLMargins
{
    public Double Left { get; set; }
    public Double Right { get; set; }
    public Double Top { get; set; }
    public Double Bottom { get; set; }
    public Double Header { get; set; }
    public Double Footer { get; set; }

    public IXLMargins SetLeft(Double value)
    {
        this.Left = value;
        return this;
    }

    public IXLMargins SetRight(Double value)
    {
        this.Right = value;
        return this;
    }

    public IXLMargins SetTop(Double value)
    {
        this.Top = value;
        return this;
    }

    public IXLMargins SetBottom(Double value)
    {
        this.Bottom = value;
        return this;
    }

    public IXLMargins SetHeader(Double value)
    {
        this.Header = value;
        return this;
    }

    public IXLMargins SetFooter(Double value)
    {
        this.Footer = value;
        return this;
    }
}

#nullable disable

namespace XlsxSharp.Excel.PageSetup;

public interface IXLMargins
{
    /// <summary>Gets or sets the Left margin.</summary>
    /// <value>The Left margin.</value>
    public double Left { get; set; }

    /// <summary>Gets or sets the Right margin.</summary>
    /// <value>The Right margin.</value>
    public double Right { get; set; }

    /// <summary>Gets or sets the Top margin.</summary>
    /// <value>The Top margin.</value>
    public double Top { get; set; }

    /// <summary>Gets or sets the Bottom margin.</summary>
    /// <value>The Bottom margin.</value>
    public double Bottom { get; set; }

    /// <summary>Gets or sets the Header margin.</summary>
    /// <value>The Header margin.</value>
    public double Header { get; set; }

    /// <summary>Gets or sets the Footer margin.</summary>
    /// <value>The Footer margin.</value>
    public double Footer { get; set; }

    public IXLMargins SetLeft(double value);
    public IXLMargins SetRight(double value);
    public IXLMargins SetTop(double value);
    public IXLMargins SetBottom(double value);
    public IXLMargins SetHeader(double value);
    public IXLMargins SetFooter(double value);
}

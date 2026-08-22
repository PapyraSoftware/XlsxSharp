namespace XlsxSharp.Excel;

public partial class XLColor
{
    internal XLColorKey Key { get; }

    private XLColor()
        : this(new XLColorKey() { ColorType = XLColorType.Automatic }) => this.HasValue = false;

    internal XLColor(XLColorKey key)
    {
        this.Key = key;
        this.HasValue = true;
    }

    /// <summary>
    /// Lower case color type for exception messages.
    /// </summary>
    private string LcColorType => this.ColorType.ToString().ToLowerInvariant();
}

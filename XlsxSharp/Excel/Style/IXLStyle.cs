#nullable disable

using System;

namespace XlsxSharp.Excel;

public interface IXLStyle : IEquatable<IXLStyle>
{
    public IXLAlignment Alignment { get; set; }

    public IXLBorder Border { get; set; }

    public IXLNumberFormat DateFormat { get; }

    public IXLFill Fill { get; set; }

    public IXLFont Font { get; set; }

    /// <summary>
    /// Should the text values of a cell saved to the file be prefixed by a quote (<c>'</c>) character?
    /// Has no effect if cell values is not a <see cref="XLDataType.Text"/>. Doesn't affect values during runtime,
    /// text values are returned without quote.
    /// </summary>
    public bool IncludeQuotePrefix { get; set; }

    public IXLNumberFormat NumberFormat { get; set; }

    public IXLProtection Protection { get; set; }

    public IXLStyle SetIncludeQuotePrefix(bool includeQuotePrefix = true);
}

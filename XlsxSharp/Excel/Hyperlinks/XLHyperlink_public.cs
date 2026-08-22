#nullable disable

using System;
using XlsxSharp.Extensions;

namespace XlsxSharp.Excel;

public partial class XLHyperlink
{
    private Uri _externalAddress;
    private String _internalAddress;

    public XLHyperlink(String address) => this.SetValues(address, String.Empty);

    public XLHyperlink(String address, String tooltip) => this.SetValues(address, tooltip);

    public XLHyperlink(IXLCell cell) => this.SetValues(cell, String.Empty);

    public XLHyperlink(IXLCell cell, String tooltip) => this.SetValues(cell, tooltip);

    public XLHyperlink(IXLRangeBase range) => this.SetValues(range, String.Empty);

    public XLHyperlink(IXLRangeBase range, String tooltip) => this.SetValues(range, tooltip);

    public XLHyperlink(Uri uri) => this.SetValues(uri, String.Empty);

    public XLHyperlink(Uri uri, String tooltip) => this.SetValues(uri, tooltip);

    public Boolean IsExternal { get; set; }

    public Uri ExternalAddress
    {
        get => this.IsExternal ? this._externalAddress : null;
        set
        {
            this._externalAddress = value;
            this.IsExternal = true;
        }
    }

#nullable enable
    /// <summary>
    /// Gets top left cell of a hyperlink range. Return <c>null</c>,
    /// if the hyperlink isn't in a worksheet.
    /// </summary>
    public IXLCell? Cell
    {
        get
        {
            if (this.Container is null)
            {
                return null;
            }

            return this.Container.GetCell(this);
        }
    }

#nullable disable

    public String InternalAddress
    {
        get
        {
            if (this.IsExternal)
            {
                return null;
            }

            if (this._internalAddress.Contains('!'))
            {
                return this._internalAddress[0] != '\''
                    ? String.Concat(
                        this._internalAddress.Substring(0, this._internalAddress.IndexOf('!'))
                            .EscapeSheetName(),
                        '!',
                        this._internalAddress.Substring(this._internalAddress.IndexOf('!') + 1)
                    )
                    : this._internalAddress;
            }

            if (this.Container is null)
            {
                throw new InvalidOperationException("Hyperlink is not attached to a worksheet.");
            }

            string sheetName = this.Container.WorksheetName;
            return String.Concat(sheetName.EscapeSheetName(), '!', this._internalAddress);
        }
        set
        {
            this._internalAddress = value;
            this.IsExternal = false;
        }
    }

    /// <summary>
    /// Tooltip displayed when user hovers over the hyperlink range. If not specified,
    /// the link target is displayed in the tooltip.
    /// </summary>
    public String Tooltip { get; set; }

    /// <inheritdoc cref="IXLHyperlinks.Delete(XLHyperlink)"/>
    public void Delete() => this.Container?.Delete(this);
}

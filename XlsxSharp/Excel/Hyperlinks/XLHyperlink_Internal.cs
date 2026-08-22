#nullable disable

using System;

namespace XlsxSharp.Excel;

public partial class XLHyperlink
{
    internal XLHyperlink() { }

    internal XLHyperlink(XLHyperlink hyperlink)
    {
        this._externalAddress = hyperlink._externalAddress;
        this._internalAddress = hyperlink._internalAddress;
        this.Tooltip = hyperlink.Tooltip;
        this.IsExternal = hyperlink.IsExternal;
    }

    internal void SetValues(string address, string tooltip)
    {
        this.Tooltip = tooltip;
        if (address[0] == '.')
        {
            this._externalAddress = new Uri(address, UriKind.Relative);
            this.IsExternal = true;
        }
        else
        {
            if (Uri.TryCreate(address, UriKind.Absolute, out Uri uri))
            {
                this._externalAddress = uri;
                this.IsExternal = true;
            }
            else
            {
                this._internalAddress = address;
                this.IsExternal = false;
            }
        }
    }

    internal void SetValues(Uri uri, string tooltip)
    {
        this.Tooltip = tooltip;
        this._externalAddress = uri;
        this.IsExternal = true;
    }

    internal void SetValues(IXLCell cell, string tooltip)
    {
        this.Tooltip = tooltip;
        this._internalAddress = cell.Address.ToString(XLReferenceStyle.A1, true);
        this.IsExternal = false;
    }

    internal void SetValues(IXLRangeBase range, string tooltip)
    {
        this.Tooltip = tooltip;
        this._internalAddress = range.RangeAddress.ToString(XLReferenceStyle.A1, true);
        this.IsExternal = false;
    }

    internal XLHyperlinks Container { get; set; }
}
